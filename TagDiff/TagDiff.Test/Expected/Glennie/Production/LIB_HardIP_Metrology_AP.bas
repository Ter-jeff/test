Attribute VB_Name = "LIB_HardIP_Metrology_AP"
'Option Explicit 'Add ErrHandler 2023/05/29
'Type MTRSNS_Matrix
'    ROT_Matrix() As Double
'    ROV_Matrix() As Double
'    ROT_a_max_min_Matrix() As Double
'    ROV_a_max_min_Matrix() As Double
'End Type
'Public MetrologySense_Matrix() As MTRSNS_Matrix
Public Flag_TMPS_1st_Run As Boolean

Public Function MetrologyTMPS_Measurement_Process(pat As String, srcPin As PinList, code() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef Rtn_MeasVolt() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName() As String, MeasV_WaitTime As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim srcWave() As New DSPWave: ReDim srcWave(UBound(code))
    Dim site As Variant
    Dim InDSPwave As New DSPWave
    Dim i As Long, j As Long
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array() As Long: ReDim srcwave_array(TrimCodeSize - 1)
    
    ByPassTestLimit = True
    'glb_Disable_CurrRangeSetting_Print = True

    For i = 0 To UBound(code)
    srcWave(i).CreateConstant 0, TrimCodeSize, DspLong
        For Each site In TheExec.sites
            For j = 0 To TrimCodeSize - 1
                If j = 0 Then
                    srcwave_array(j) = code(i) And 1
                Else
                    srcwave_array(j) = (code(i) And (2 ^ j)) \ (2 ^ j)
                End If
            Next j
        srcWave(i).data = srcwave_array
        Next site
        Call StoreDataAllType(TrimStoreName(i), srcWave(i))
    Next i
    
    Call GeneralDigSrcSettingWithBurst(pat, srcPin, InDSPwave)

    TheHdw.patterns(pat).start
    
    For i = 0 To NumberOfMeasV - 1
        TheHdw.Digital.Patgen.FlagWait cpuA, 0
        Rtn_MeasVolt(i) = HardIP_MeasureVolt
        Call DebugPrintFunc_PPMU(vbNullString)
        
        'Oscar, DiffMeter
                
        For Each site In TheExec.sites
            Res(i) = Abs(Rtn_MeasVolt(i).pins(0).value - Rtn_MeasVolt(i).pins(1).value)
                        'Res(i) = Abs(Rtn_MeasVolt(i).Pins(0).value)
            For j = 0 To Rtn_MeasVolt(i).pins.Count - 1
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i) & ", Pin : " & Rtn_MeasVolt(i).pins(j) & ", Voltage = " & Rtn_MeasVolt(i).pins(j).value
            Next j
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i) & ", Voltage Difference = " & Res(i)
        Next site
        TheHdw.Digital.Patgen.Continue 0, cpuA
    Next i
    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    'glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "MetrologyTMPS_Measurement_Process") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MTRSNS_IDLoading(MTX_ID_FIX As Long, MTR_Version As Long, MTX_ID_VAR As Long, MTR_T1_Version As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim Max_Rows_Count As Long
Dim Max_Columns_Count As Long
Dim MTRSNS_Matrix_Range As Variant
Dim MTRSNS_Matrix_Sheet As Worksheet: Set MTRSNS_Matrix_Sheet = Sheets("MTRSNS_Matrix")
Dim i As Integer: i = 0
'Dim MTX_ID_FIX As Long
'Dim MTR_Version As Long
Dim First_ID As Long
'Dim MTX_ID_VAR As Long
'Dim MTR_T1_Version As Long

With MTRSNS_Matrix_Sheet
    Max_Rows_Count = .UsedRange.Rows.Count
    Max_Columns_Count = .UsedRange.Columns.Count
    MTRSNS_Matrix_Range = .range(.Cells(1, 1), .Cells(Max_Rows_Count, Max_Columns_Count))
End With



For i = 1 To Max_Rows_Count
    'Debug.Print MTRSNS_Matrix_Sheet.Cells(i, 1).Value
    If MTRSNS_Matrix_Sheet.Cells(i, 1).value Like "*Matrix ID*" Then
       MTX_ID_FIX = MTRSNS_Matrix_Sheet.Cells(i, 4).value
       MTR_Version = MTRSNS_Matrix_Sheet.Cells(i + 1, 4).value
       First_ID = i + 1
       Exit For
    End If
    
Next i


For i = First_ID To Max_Rows_Count

    If MTRSNS_Matrix_Sheet.Cells(i, 1).value Like "*Matrix ID*" Then
       MTX_ID_VAR = MTRSNS_Matrix_Sheet.Cells(i, 4).value
       MTR_T1_Version = MTRSNS_Matrix_Sheet.Cells(i + 1, 4).value
       First_ID = i + 1
       Exit For
    End If

Next i

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "MTRSNS_IDLoading") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MetrologyTMPS_OffSet(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
Dim DSP_OffSet_Mean As New DSPWave
Dim DSP_OffSet_Mean_Array(0) As Double
Dim DSP_OffSet_Mean_eFuse As New DSPWave
'Dim DSP_OffSet_Mean_Fuse_Array(0) As Double
Dim TestNameInput As String
Dim i As Long
    For i = 0 To UBound(InWf_Array)
        InWf_SiteDouble(i) = GetStoreDataAllType(InWf_Split(i) & "_para")
    Next i
    For Each site In TheExec.sites.Active
        For i = 0 To UBound(InWf_Array)
            InWf_Array(i) = InWf_SiteDouble(i)
        Next i
        InWf.data = InWf_Array
        DSP_OffSet_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0)
        DSP_OffSet_Mean.data = DSP_OffSet_Mean_Array
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=DSP_OffSet_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
    Call StoreDataAllType(argv(1) & "_" & CStr(TheExec.flow.var(argv(2)).value), DSP_OffSet_Mean)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_OffSet_Mean_eFuse, DSP_OffSet_Mean, 18, 0)
    Call StoreDataAllType(argv(1) & "_eFuse_" & CStr(TheExec.flow.var(argv(2)).value), DSP_OffSet_Mean_eFuse)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyTMPS_OffSet") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyTMPS_OffSet_B0TTR(argc As Integer, argv() As String) As Long
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
Dim DSP_OffSet_Mean As New DSPWave
Dim DSP_OffSet_Mean_Array(0) As Double
Dim DSP_OffSet_Mean_eFuse As New DSPWave
'Dim DSP_OffSet_Mean_Fuse_Array(0) As Double
Dim TestNameInput As String
Dim i As Long
    For i = 0 To UBound(InWf_Array)
        InWf_SiteDouble(i) = GetStoredData(InWf_Split(i) & "_para")
    Next i
    For Each site In TheExec.sites.Active
        For i = 0 To UBound(InWf_Array)
            InWf_Array(i) = InWf_SiteDouble(i)
        Next i
        InWf.data = InWf_Array
        DSP_OffSet_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0)
        DSP_OffSet_Mean.data = DSP_OffSet_Mean_Array
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=DSP_OffSet_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
    Call AddStoredCaptureData(argv(1) & "_" & "15", DSP_OffSet_Mean)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_OffSet_Mean_eFuse, DSP_OffSet_Mean, 18, 0)
    Call AddStoredCaptureData(argv(1) & "_eFuse_" & "15", DSP_OffSet_Mean_eFuse)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyTMPS_OffSet"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologyTMPS_Gain(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")

Dim DSP_Gain_Mean As New DSPWave
Dim DSP_Gain_Mean_Array(0) As Double
Dim DSP_OffSet_Mean As New DSPWave
Dim DSP_OffSet_Mean_Array() As Double
Dim DSP_Gain_Mean_Final As New DSPWave
Dim DSP_Gain_Mean_Final_Array(0) As Double
Dim TestNameInput As String
Dim i As Long

For i = 0 To UBound(InWf_Array)
    InWf_SiteDouble(i) = GetStoreDataAllType(InWf_Split(i) & "_para")
Next i

DSP_OffSet_Mean = GetStoreDataAllType(argv(1) & "_" & CStr(TheExec.flow.var(argv(2)).value))
For Each site In TheExec.sites.Active
    DSP_OffSet_Mean_Array = DSP_OffSet_Mean.data
    For i = 0 To UBound(InWf_Array)
        InWf_Array(i) = InWf_SiteDouble(i)
    Next i
    InWf.data = InWf_Array
    DSP_Gain_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0) - DSP_OffSet_Mean_Array(0)
    If DSP_Gain_Mean_Array(0) < 0 Then: DSP_Gain_Mean_Array(0) = 0
    DSP_Gain_Mean.data = DSP_Gain_Mean_Array
Next site

TestNameInput = Report_TName_From_Instance("CalcC", "X", Replace(argv(3), "_", vbNullString), 0, 0)
TheExec.flow.TestLimit resultVal:=DSP_Gain_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

Call StoreDataAllType(argv(3) & "_" & CStr(TheExec.flow.var(argv(2)).value), DSP_Gain_Mean)

If TheExec.flow.var(argv(2)).value = argv(5) Then
    For Each site In TheExec.sites.Active
        InWf.Clear
        For i = argv(4) To argv(5)
            If i = argv(4) Then
                InWf = GetStoreDataAllType(argv(3) & "_" & i)
            Else
                InWf = InWf.Concatenate(GetStoreDataAllType(argv(3) & "_" & i))
            End If
        Next i
        
    DSP_Gain_Mean_Final_Array(0) = FormatNumber(InWf.CalcMean, 0)
    DSP_Gain_Mean_Final.data = DSP_Gain_Mean_Final_Array
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_AVG", "_", vbNullString))
    TheExec.flow.TestLimit resultVal:=DSP_Gain_Mean_Final.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    Call StoreDataAllType(argv(3), DSP_Gain_Mean_Final)
End If

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyTMPS_Gain") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MTRTMPS_DSSCOUT_AVG(InWf As DSPWave, StoreName_DSSC_OUT_Mean As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim DSP_DSSCOUT_Mean As New DSPWave
Dim DSP_DSSCOUT_Mean_Array(0) As Double
Dim TestNameInput As String

For Each site In TheExec.sites.Active
    DSP_DSSCOUT_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0)
    DSP_DSSCOUT_Mean.data = DSP_DSSCOUT_Mean_Array
Next site

TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
TheExec.flow.TestLimit resultVal:=DSP_DSSCOUT_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
Call StoreDataAllType(StoreName_DSSC_OUT_Mean, DSP_DSSCOUT_Mean)

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "MTRTMPS_DSSCOUT_AVG") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MetrologyGR_Measurement_Process(pat As String, srcPin As PinList, code As SiteLong, Res As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef Rtn_MeasVolt() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName As String, MeasV_WaitTime As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim sigName As String, srcWave As New DSPWave, site As Variant
    Dim InDSPwave As New DSPWave
    Dim i As Long, j As Long
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array() As Long: ReDim srcwave_array(TrimCodeSize - 1)
    srcWave.CreateConstant 0, TrimCodeSize, DspLong
    ByPassTestLimit = True
    'glb_Disable_CurrRangeSetting_Print = True

    For Each site In TheExec.sites
        For i = 0 To TrimCodeSize - 1
            If i = 0 Then
                srcwave_array(i) = code And 1
            Else
                srcwave_array(i) = (code And (2 ^ i)) \ (2 ^ i)
            End If
        Next i
    srcWave.data = srcwave_array
    Next site
    
    Call StoreDataAllType(TrimStoreName, srcWave)
    
    Call GeneralDigSrcSettingWithBurst(LCase(pat), srcPin, InDSPwave)
   
    TheHdw.patterns(pat).start
    
    For i = 0 To NumberOfMeasV - 1
        Instance_Data.TestSeqNum = i
        TheHdw.Digital.Patgen.FlagWait cpuA, 0

        Rtn_MeasVolt(i) = HardIP_MeasureVolt
        Call DebugPrintFunc_PPMU(vbNullString)

        For Each site In TheExec.sites

            If i = NumberOfMeasV - 1 Then Res = Abs(Rtn_MeasVolt(i).pins(0).value - Rtn_MeasVolt(i - 1).pins(0).value)
            For j = 0 To Rtn_MeasVolt(i).pins.Count - 1
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Pin : " & Rtn_MeasVolt(i).pins(j) & ", Voltage = " & Rtn_MeasVolt(i).pins(j).value
            Next j
            If gl_Disable_HIP_debug_log = False Then If i = NumberOfMeasV - 1 Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Voltage Difference = " & Res
        Next site

        TheHdw.Digital.Patgen.Continue 0, cpuA
    Next i
    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    'glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "MetrologyGR_Measurement_Process") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyGR_TRC(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim DSP_APM_Count_Out_0 As New DSPWave: DSP_APM_Count_Out_0 = GetStoreDataAllType(argv(0))
    Dim DSP_APM_Count_Out_1 As New DSPWave: DSP_APM_Count_Out_1 = GetStoreDataAllType(argv(1))
    Dim DictionaryName As String: DictionaryName = argv(2)
    Dim DSP_T1 As New DSPWave
    Dim DSP_T2 As New DSPWave
    Dim DSP_TRC As New DSPWave
    Dim DSP_TRC_ERR As New DSPWave
    Dim DSP_TRC_ERR_eFuse As New DSPWave
    Dim TestNameInput As String
    Dim site As Variant
    For Each site In TheExec.sites
        DSP_T1 = DSP_APM_Count_Out_0.ConvertStreamTo(tldspParallel, DSP_APM_Count_Out_0.SampleSize, 0, Bit0IsMsb).Multiply(375000).Add(0.001).Reciprocate
        DSP_T2 = DSP_APM_Count_Out_1.ConvertStreamTo(tldspParallel, DSP_APM_Count_Out_1.SampleSize, 0, Bit0IsMsb).Multiply(375000).Add(0.001).Reciprocate
        DSP_TRC = DSP_T2.Subtract(DSP_T1)
        DSP_TRC_ERR = DSP_TRC.Subtract(0.0000000017).divide(0.0000000017).Multiply(100).divide(0.5)
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_T1.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_T2.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_TRC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_TRC_ERR.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_TRC_ERR_eFuse, DSP_TRC_ERR, 8, 0)
    Call StoreDataAllType(DictionaryName, DSP_TRC_ERR_eFuse)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyGR_TRC") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologySense_Frequency(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim TestNameInput As String
    Dim site As Variant
    Dim MetrologySense_Frequency() As New DSPWave: ReDim MetrologySense_Frequency(argc - 1)
    Dim SubBlockName As String: SubBlockName = Split(TheExec.DataManager.instancename, "_")(1)
    For i = 0 To argc - 1
        For Each site In TheExec.sites
            DSPWave_Dict = GetStoreDataAllType(argv(i))
            MetrologySense_Frequency(i) = DSPWave_Dict.ConvertStreamTo(tldspParallel, 21, 0, Bit0IsMsb).Multiply(50000)
        Next
        Call StoreDataAllType(SubBlockName & "-" & Instance_Data.Tname(TheExec.flow.TestLimitIndex), MetrologySense_Frequency(i))
        TestNameInput = Report_TName_From_Instance("CalcF", vbNullString)
        TheExec.flow.TestLimit resultVal:=MetrologySense_Frequency(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
If UCase(TheExec.DataManager.instancename) = "MTRSNS_ASGMTRT1P3VDDCPUSRAMV1P250VDDPCPUV0P750VDDECPUV0P750_PP_SCYA0_C_FULP_AN_MT03_DLL_JTG_COD_ALLFV_SI_ASGMTR_T1P3_LV" Then Flag_TMPS_1st_Run = False
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologySense_Frequency") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyTMPS_Temperature(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature_Average As New DSPWave
    
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    DSP_MTRSNS_Temperature_Average.CreateConstant 0, UBound(Temperature_Sensor) + 1
    
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).data = Temperature_Array
            DSP_MTRSNS_Temperature_Average.Element(i) = DSP_MTRSNS_Temperature_Average.Element(i) + DSP_Temperature(i).Element(0)
        Next site
    Next i
    If argc = 2 And InStr(argv(1), "avg") = 0 Then
        Temperature_Dictionary = Split(argv(1), "+")
        For i = 0 To UBound(Temperature_Sensor)
            For Each site In TheExec.sites
                DSP_MTRSNS_Temperature(i) = DSP_Temperature(i).Multiply(8)
            Next site
'            If UCase(theexec.CurrentJob) = "WLFT1" Then
'                Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_MTRSNS_Temperature_eFuse(i), DSP_MTRSNS_Temperature(i), 10, 0)
'            Else
                Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_MTRSNS_Temperature_eFuse(i), DSP_MTRSNS_Temperature(i), 11, 0)
'            End If
            Call StoreDataAllType(Temperature_Dictionary(i), DSP_MTRSNS_Temperature_eFuse(i))
        Next i
        
    Else
    
        If InStr(argv(1), "avg") <> 0 Then
            For Each site In TheExec.sites
                DSP_MTRSNS_Temperature_Average(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Average(site).CalcMean, 0)
            Next site
            
            If UCase(TheExec.DataManager.instancename) Like "*TFEVERIF_T3P3_NV" Then
                Call AddStoredCaptureData(argv(1), DSP_MTRSNS_Temperature_Average)
            End If
        End If
        
    End If
    
    For i = 0 To UBound(Temperature_Sensor)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
    
     If argc = 2 And InStr(argv(1), "avg") <> 0 Then
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
      
     End If
     
    
'    If TheExec.Flow.EnableWord("TMPS_Monitor") = True Then
'        TheHdw.Pins("all_digital").Digital.InitState = chInitLo
'        TheHdw.DCVS.Pins("VDD_SRAM_GPU,VDD_GPU,VDD_ECPU,VDD_PCPU,VDD_CPU_SRAM").Voltage.Main.value = 0.5
'        If Not (Flag_TMPS_1st_Run) Then
'            TheHdw.Wait 5
'            Flag_TMPS_1st_Run = True
'        Else
'            TheHdw.Wait 1
'        End If
'    End If
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyTMPS_Temperature") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyTMPS_Vref_Error(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim Str_Vref As String: Str_Vref = argv(0)
    Dim Fuse_BitCount As Long: Fuse_BitCount = CLng(argv(1))
    
    Dim SiteDbl_Vref As New SiteDouble: SiteDbl_Vref = GetStoreDataAllType(Str_Vref)
    Dim SiteDbl_Vref_Error As New SiteDouble: SiteDbl_Vref_Error = SiteDbl_Vref.divide(0.8).Subtract(1).Multiply(100).divide(0.125)
    Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Fuse_BitCount - 1, "1"))
    Dim Low_limit As Double: Low_limit = -2 ^ (Fuse_BitCount - 1)
    Dim site As Variant
    For Each site In TheExec.sites
        SiteDbl_Vref_Error = Floor(SiteDbl_Vref_Error)
    Next site
    Dim TestNameInput As String: TestNameInput = Report_TName_From_Instance("CalcC", "X", "vreferr", 0, 0)
    
    TheExec.flow.TestLimit resultVal:=SiteDbl_Vref_Error, lowVal:=Low_limit, hiVal:=High_limit, Tname:=TestNameInput, ForceResults:=tlForceFlow

    Dim SiteDbl_Vref_Error_Fuse As New DSPWave
    Dim SiteDbl_Vref_Error_Fuse_Array(0) As Long

    For Each site In TheExec.sites
        SiteDbl_Vref_Error = FormatNumber(SiteDbl_Vref_Error, 0)
        If SiteDbl_Vref_Error < Low_limit Then
            SiteDbl_Vref_Error_Fuse_Array(0) = 2 ^ (Fuse_BitCount) + FormatNumber(Low_limit)
        ElseIf SiteDbl_Vref_Error >= Low_limit And SiteDbl_Vref_Error < 0 Then
            SiteDbl_Vref_Error_Fuse_Array(0) = 2 ^ (Fuse_BitCount) + FormatNumber(SiteDbl_Vref_Error)
        ElseIf SiteDbl_Vref_Error < High_limit And SiteDbl_Vref_Error >= 0 Then
            SiteDbl_Vref_Error_Fuse_Array(0) = FormatNumber(SiteDbl_Vref_Error)
        Else
            SiteDbl_Vref_Error_Fuse_Array(0) = FormatNumber(High_limit)
        End If
        SiteDbl_Vref_Error_Fuse.data = SiteDbl_Vref_Error_Fuse_Array
    Next site
    
    Call StoreDataAllType(argv(argc - 1), SiteDbl_Vref_Error_Fuse)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyTMPS_Vref_Error") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MetrologyTMPS_Coefficients(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim InWf As New DSPWave
    Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
    Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
    Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
    Dim TempVal As Long
    Dim DSP_Offset As New DSPWave: DSP_Offset = GetStoreDataAllType(argv(1))
    Dim DSP_Gain As New DSPWave: DSP_Gain = GetStoreDataAllType(argv(2))
    Dim Integer_Bit As Long: Integer_Bit = argv(4)
    Dim Fractional_Bit As Long: Fractional_Bit = argv(5)
    Dim DSP_ADC_Temperature_Sensor_Raw_Data As New DSPWave
    Dim ADC_Temperature_Sensor_Raw_Data_Array(0) As Double
    Dim DSP_V0 As New DSPWave
    Dim DSP_V1 As New DSPWave
    Dim DSP_X0A As New DSPWave
    Dim DSP_X0 As New DSPWave
    Dim DSP_a0Cal As New DSPWave
    Dim a0 As Double
    Dim a1 As Double
    Dim a2 As Double
    Dim a3 As Double
    Dim DSP_Coefficient_C0 As New DSPWave
    Dim DSP_Coefficient_C1 As New DSPWave
    Dim DSP_Coefficient_C2 As New DSPWave
    Dim DSP_Coefficient_C3 As New DSPWave
    Dim DSP_Coefficient_C0_eFuse As New DSPWave
    Dim DSP_Coefficient_C1_eFuse As New DSPWave
    Dim DSP_Coefficient_C2_eFuse As New DSPWave
    Dim DSP_Coefficient_C3_eFuse As New DSPWave
    Dim Array_Coefficient_C0_eFuse() As Double
    Dim Array_Coefficient_C1_eFuse() As Double
    Dim Array_Coefficient_C2_eFuse() As Double
    Dim Array_Coefficient_C3_eFuse() As Double
    Dim DSP_Coefficient_C0_Source As New DSPWave
    Dim DSP_Coefficient_C1_Source As New DSPWave
    Dim DSP_Coefficient_C2_Source As New DSPWave
    Dim DSP_Coefficient_C3_Source As New DSPWave
    Dim Array_Coefficient_C0_Source() As Long: ReDim Array_Coefficient_C0_Source(Integer_Bit + Fractional_Bit - 1)
    Dim Array_Coefficient_C1_Source() As Long: ReDim Array_Coefficient_C1_Source(Integer_Bit + Fractional_Bit - 1)
    Dim Array_Coefficient_C2_Source() As Long: ReDim Array_Coefficient_C2_Source(Integer_Bit + Fractional_Bit - 1)
    Dim Array_Coefficient_C3_Source() As Long: ReDim Array_Coefficient_C3_Source(Integer_Bit + Fractional_Bit - 1)
    Dim BKM_DECODE As String
    Dim i As Long
    Dim TestNameInput As String
    For i = 0 To UBound(InWf_Array)
        InWf_SiteDouble(i) = GetStoreDataAllType(InWf_Split(i) & "_para")
    Next i
    For Each site In TheExec.sites.Active
        For i = 0 To UBound(InWf_Array)
            InWf_Array(i) = InWf_SiteDouble(i)
        Next i
        InWf.data = InWf_Array
        ADC_Temperature_Sensor_Raw_Data_Array(0) = FormatNumber(InWf.CalcMean, 0)
        DSP_ADC_Temperature_Sensor_Raw_Data.data = ADC_Temperature_Sensor_Raw_Data_Array
    Next site
    For Each site In TheExec.sites
        BKM_DECODE = gS_BKM_IEDA
        Exit For
    Next site

        If TheExec.TesterMode = testModeOffline Then BKM_DECODE = 2

    If BKM_DECODE = "1" Then
        a0 = -179.0341695
        a1 = 679.7734778
        a2 = -288.9625211
        a3 = 50.6003026
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM1.8 ) **********************"
        TheExec.Datalog.WriteComment "  a0 = -179.0341695(0xFE99EE8)"
        TheExec.Datalog.WriteComment "  a1 = 679.7734778(0x54F8C0)"
        TheExec.Datalog.WriteComment "  a2 = -288.9625211(0xFDBE133)"
        TheExec.Datalog.WriteComment "  a3 = 50.6003026(0x65335)"
        TheExec.Datalog.WriteComment "*************************************************************************************"
    ElseIf BKM_DECODE = "2" Then
        a0 = -208.7079186
        a1 = 762.9699164
        a2 = -355.6894227
        a3 = 65.09103582
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM4.0 ) **********************"
        TheExec.Datalog.WriteComment "  a0 = -208.7079186(0xFE5E959)"
        TheExec.Datalog.WriteComment "  a1 = 762.9699164(0x5F5F0A)"
        TheExec.Datalog.WriteComment "  a2 = -355.6894227(0xFD389F0)"
        TheExec.Datalog.WriteComment "  a3 = 65.09103582(0x822EA)"
        TheExec.Datalog.WriteComment "*************************************************************************************"
    ElseIf BKM_DECODE = "3" Then
        a0 = -208.7079186
        a1 = 762.9699164
        a2 = -355.6894227
        a3 = 65.09103582
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM4.2 ) **********************"
        TheExec.Datalog.WriteComment "  a0 = -208.7079186(0xFE5E959)"
        TheExec.Datalog.WriteComment "  a1 = 762.9699164(0x5F5F0A)"
        TheExec.Datalog.WriteComment "  a2 = -355.6894227(0xFD389F0)"
        TheExec.Datalog.WriteComment "  a3 = 65.09103582(0x822EA)"
        TheExec.Datalog.WriteComment "*************************************************************************************"
    ElseIf BKM_DECODE = "4" Then
        a0 = -208.7079186
        a1 = 762.9699164
        a2 = -355.6894227
        a3 = 65.09103582
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM4.5 ) **********************"
        TheExec.Datalog.WriteComment "  a0 = -208.7079186(0xFE5E959)"
        TheExec.Datalog.WriteComment "  a1 = 762.9699164(0x5F5F0A)"
        TheExec.Datalog.WriteComment "  a2 = -355.6894227(0xFD389F0)"
        TheExec.Datalog.WriteComment "  a3 = 65.09103582(0x822EA)"
        TheExec.Datalog.WriteComment "*************************************************************************************"
    ElseIf BKM_DECODE = "5" Then
        a0 = -208.7079186
        a1 = 762.9699164
        a2 = -355.6894227
        a3 = 65.09103582
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM5.0 ) **********************"
        TheExec.Datalog.WriteComment "  a0 = -208.7079186(0xFE5E959)"
        TheExec.Datalog.WriteComment "  a1 = 762.9699164(0x5F5F0A)"
        TheExec.Datalog.WriteComment "  a2 = -355.6894227(0xFD389F0)"
        TheExec.Datalog.WriteComment "  a3 = 65.09103582(0x822EA)"
        TheExec.Datalog.WriteComment "*************************************************************************************"
    ElseIf BKM_DECODE = "6" Then
        a0 = -59.285102
        a1 = 504.03940702
        a2 = -186.08443514
        a3 = 29.81002933
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM5.01 ) **********************"
        TheExec.Datalog.WriteComment "  a0 = -59.285102 (0x76920)"
        TheExec.Datalog.WriteComment "  a1 = 504.03940702 (0x3F0143)"
        TheExec.Datalog.WriteComment "  a2 = -186.08443514 (0xFE8BD4C)"
        TheExec.Datalog.WriteComment "  a3 = 29.81002933 (0x3B9EC)"
        TheExec.Datalog.WriteComment "*************************************************************************************"
    ElseIf BKM_DECODE = "7" Then
        a0 = -59.285102
        a1 = 504.03940702
        a2 = -186.08443514
        a3 = 29.81002933
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM5.11 ) **********************"
        TheExec.Datalog.WriteComment "  a0 = -59.285102 (0x76920)"
        TheExec.Datalog.WriteComment "  a1 = 504.03940702 (0x3F0143)"
        TheExec.Datalog.WriteComment "  a2 = -186.08443514 (0xFE8BD4C)"
        TheExec.Datalog.WriteComment "  a3 = 29.81002933 (0x3B9EC)"
        TheExec.Datalog.WriteComment "*************************************************************************************"
    Else
'        A0 = -239.2907296
'        a1 = 788.6170872
'        a2 = -355.6894227
'        a3 = 65.09103582
        TheExec.Datalog.WriteComment "********************** Unknown BKM **********************"
'        TheExec.Datalog.WriteComment "  a0 = -239.2907296(0xFE216B2)"
'        TheExec.Datalog.WriteComment "  a1 = 788.6170872(0x6293BF)"
'        TheExec.Datalog.WriteComment "  a2 = -355.6894227(0xFD389F0)"
'        TheExec.Datalog.WriteComment "  a3 = 65.09103582(0x822EA)"
'        TheExec.Datalog.WriteComment "*************************************************************************************"
'        DSP_Coefficient_C0.CreateConstant 0, 1, DspDouble
'        DSP_Coefficient_C1.CreateConstant 0, 1, DspDouble
'        DSP_Coefficient_C2.CreateConstant 0, 1, DspDouble
'        DSP_Coefficient_C3.CreateConstant 0, 1, DspDouble
        For Each site In TheExec.sites
            TheExec.sites.item(site).FlagState("F_HardIP_MTRTSNS_Unknown_BKM_Flag") = logicTrue
        Next site
        Exit Function
    End If

    For Each site In TheExec.sites
        DSP_V0 = DSP_Offset.divide(2 ^ 13)
        DSP_V1 = DSP_Gain.divide(2 ^ 14)
        DSP_X0A = DSP_ADC_Temperature_Sensor_Raw_Data.divide(2 ^ 13)
        DSP_X0 = DSP_X0A.Subtract(DSP_V0).divide(DSP_V1.Add(0.0000000001)) 'x0=(x0a-v0)/v1;
        DSP_a0Cal = DSP_X0.Multiply(-a1).Add(DSP_X0.Square.Multiply(-a2)).Add(DSP_X0.Square.Multiply(DSP_X0).Multiply(-a3)).Add(25).Add(273.15) 'a0cal=25+273.15-a1*x0-a2*x0^2-a3*x0^3
        DSP_Coefficient_C0 = DSP_a0Cal.Subtract(DSP_V0.divide(DSP_V1.Add(0.0000000001)).Multiply(a1)) 'coef_c0=a0cal-a1*v0/v1;
        DSP_Coefficient_C1 = DSP_V1.Add(0.0000000001).Reciprocate.Multiply(a1).Subtract(DSP_V0.divide(DSP_V1.Add(0.0000000001).Square).Multiply(2).Multiply(a2)) 'coef_c1=a1/v1-2*a2*v0/v1^2;
        DSP_Coefficient_C2 = DSP_V1.Add(0.0000000001).Square.Reciprocate.Multiply(a2).Subtract(DSP_V0.divide(DSP_V1.Add(0.0000000001).Square.Multiply(DSP_V1.Add(0.0000000001))).Multiply(3).Multiply(a3)) 'coef_c2=a2/v1^2-3*a3*v0/v1^3;
        DSP_Coefficient_C3 = DSP_V1.Add(0.0000000001).Square.Multiply(DSP_V1.Add(0.0000000001)).Reciprocate.Multiply(a3) 'coef_c3=a3/v1^3
    Next site

    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_Coefficient_C0_eFuse, DSP_Coefficient_C0, Integer_Bit, Fractional_Bit)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_Coefficient_C1_eFuse, DSP_Coefficient_C1, Integer_Bit, Fractional_Bit)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_Coefficient_C2_eFuse, DSP_Coefficient_C2, Integer_Bit, Fractional_Bit)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_Coefficient_C3_eFuse, DSP_Coefficient_C3, Integer_Bit, Fractional_Bit)
    For Each site In TheExec.sites
        Array_Coefficient_C0_eFuse = DSP_Coefficient_C0_eFuse.data
        TempVal = Array_Coefficient_C0_eFuse(0)
            For i = 0 To Integer_Bit + Fractional_Bit - 1
                Array_Coefficient_C0_Source(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        DSP_Coefficient_C0_Source.data = Array_Coefficient_C0_Source
        Array_Coefficient_C1_eFuse = DSP_Coefficient_C1_eFuse.data
        TempVal = Array_Coefficient_C1_eFuse(0)
            For i = 0 To Integer_Bit + Fractional_Bit - 1
                Array_Coefficient_C1_Source(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        DSP_Coefficient_C1_Source.data = Array_Coefficient_C1_Source
        Array_Coefficient_C2_eFuse = DSP_Coefficient_C2_eFuse.data
        TempVal = Array_Coefficient_C2_eFuse(0)
            For i = 0 To Integer_Bit + Fractional_Bit - 1
                Array_Coefficient_C2_Source(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        DSP_Coefficient_C2_Source.data = Array_Coefficient_C2_Source
        Array_Coefficient_C3_eFuse = DSP_Coefficient_C3_eFuse.data
        TempVal = Array_Coefficient_C3_eFuse(0)
            For i = 0 To Integer_Bit + Fractional_Bit - 1
                Array_Coefficient_C3_Source(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        DSP_Coefficient_C3_Source.data = Array_Coefficient_C3_Source
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=DSP_ADC_Temperature_Sensor_Raw_Data.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_coeff0", "_", vbNullString))
    TheExec.flow.TestLimit resultVal:=DSP_Coefficient_C0.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_coeff1", "_", vbNullString))
    TheExec.flow.TestLimit resultVal:=DSP_Coefficient_C1.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_coeff2", "_", vbNullString))
    TheExec.flow.TestLimit resultVal:=DSP_Coefficient_C2.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_coeff3", "_", vbNullString))
    TheExec.flow.TestLimit resultVal:=DSP_Coefficient_C3.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow

    Call StoreDataAllType(argv(3) & "_COEF_0_ATE_para", DSP_Coefficient_C0_eFuse)   'Store Decimal data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_1_ATE_para", DSP_Coefficient_C1_eFuse)   'Store Decimal data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_2_ATE_para", DSP_Coefficient_C2_eFuse)   'Store Decimal data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_3_ATE_para", DSP_Coefficient_C3_eFuse)   'Store Decimal data -- 20221115

    Call StoreDataAllType(argv(3) & "_COEF_0_ATE", DSP_Coefficient_C0_Source)               'Store Binary data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_1_ATE", DSP_Coefficient_C1_Source)               'Store Binary data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_2_ATE", DSP_Coefficient_C2_Source)               'Store Binary data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_3_ATE", DSP_Coefficient_C3_Source)               'Store Binary data -- 20221115
        
    Call StoreDataAllType(argv(3) & "_COEF_0_ATE_SRC", DSP_Coefficient_C0_Source)   'Store Binary data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_1_ATE_SRC", DSP_Coefficient_C1_Source)   'Store Binary data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_2_ATE_SRC", DSP_Coefficient_C2_Source)   'Store Binary data -- 20221115
    Call StoreDataAllType(argv(3) & "_COEF_3_ATE_SRC", DSP_Coefficient_C3_Source)   'Store Binary data -- 20221115

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyTMPS_Coefficients") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MTRSNS_Matrix_Loading()
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim MTRSNS_Matrix_Sheet As Worksheet: Set MTRSNS_Matrix_Sheet = Sheets("MTRSNS_Matrix")
Dim Column_Index As Long: Column_Index = 1
Dim Row_Index As Long: Row_Index = 1
Dim Matrix_Index As Long
Dim MetrologySense_Matrix_Index As Long: MetrologySense_Matrix_Index = 0
Dim MTRSNS_Matrix_Range As Variant
Dim Max_Rows_Count As Long
Dim Max_Columns_Count As Long

With MTRSNS_Matrix_Sheet
    Max_Rows_Count = .UsedRange.Rows.Count
    Max_Columns_Count = .UsedRange.Columns.Count
    MTRSNS_Matrix_Range = .range(.Cells(1, 1), .Cells(Max_Rows_Count, Max_Columns_Count))
End With

While Row_Index <> Max_Rows_Count
    ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index)
    If UCase(MTRSNS_Matrix_Range(Row_Index, 1)) Like "*ROT*" Then
        Row_Index = Row_Index + 1
        Matrix_Index = 0
        While MTRSNS_Matrix_Range(Row_Index, 1) <> ""
            Column_Index = 1
            Do While Column_Index <= Max_Columns_Count
                If MTRSNS_Matrix_Range(Row_Index, Column_Index) = "" Then Exit Do
                ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index).ROT_Matrix(Matrix_Index)
                MetrologySense_Matrix(MetrologySense_Matrix_Index).ROT_Matrix(Matrix_Index) = MTRSNS_Matrix_Range(Row_Index, Column_Index)
                Column_Index = Column_Index + 1
                Matrix_Index = Matrix_Index + 1
            Loop
            Row_Index = Row_Index + 1
        Wend
        Do
            Row_Index = Row_Index + 1
        Loop Until UCase(MTRSNS_Matrix_Range(Row_Index, 1)) Like "*A_MAX*"
        Row_Index = Row_Index + 1
        Matrix_Index = 0
        While MTRSNS_Matrix_Range(Row_Index, 1) <> ""
            Column_Index = 1
            Do While Column_Index <= Max_Columns_Count
                If MTRSNS_Matrix_Range(Row_Index, Column_Index) = "" Then Exit Do
                ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index).ROT_a_max_min_Matrix(Matrix_Index)
                MetrologySense_Matrix(MetrologySense_Matrix_Index).ROT_a_max_min_Matrix(Matrix_Index) = MTRSNS_Matrix_Range(Row_Index, Column_Index)
                Column_Index = Column_Index + 1
                Matrix_Index = Matrix_Index + 1
            Loop
            Row_Index = Row_Index + 1
        Wend
    ElseIf UCase(MTRSNS_Matrix_Range(Row_Index, 1)) Like "*ROV*" Then
        Row_Index = Row_Index + 1
        Matrix_Index = 0
        While MTRSNS_Matrix_Range(Row_Index, 1) <> ""
            Column_Index = 1
            Do While Column_Index <= Max_Columns_Count
                If MTRSNS_Matrix_Range(Row_Index, Column_Index) = "" Then Exit Do
                ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index).ROV_Matrix(Matrix_Index)
                MetrologySense_Matrix(MetrologySense_Matrix_Index).ROV_Matrix(Matrix_Index) = MTRSNS_Matrix_Range(Row_Index, Column_Index)
                Column_Index = Column_Index + 1
                Matrix_Index = Matrix_Index + 1
            Loop
            Row_Index = Row_Index + 1
        Wend
        Do
            Row_Index = Row_Index + 1
        Loop Until UCase(MTRSNS_Matrix_Range(Row_Index, 1)) Like "*A_MAX*"
        Row_Index = Row_Index + 1
        Matrix_Index = 0
        While MTRSNS_Matrix_Range(Row_Index, 1) <> ""
            Column_Index = 1
            Do While Column_Index <= Max_Columns_Count
                If MTRSNS_Matrix_Range(Row_Index, Column_Index) = "" Then Exit Do
                ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index).ROV_a_max_min_Matrix(Matrix_Index)
                MetrologySense_Matrix(MetrologySense_Matrix_Index).ROV_a_max_min_Matrix(Matrix_Index) = MTRSNS_Matrix_Range(Row_Index, Column_Index)
                Column_Index = Column_Index + 1
                Matrix_Index = Matrix_Index + 1
            Loop
            Row_Index = Row_Index + 1
        Wend
    ElseIf Replace(UCase(MTRSNS_Matrix_Range(Row_Index, 1)), " ", vbNullString) Like "*MATRIXID*" Then
        MetrologySense_Matrix_Index = MetrologySense_Matrix_Index + 1
    Else
    'Do nothing
    End If
    Row_Index = Row_Index + 1
Wend
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "MTRSNS_Matrix_Loading") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologySense_Compression(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
If UCase(TheExec.CurrentJob) = "CP1" Or UCase(TheExec.CurrentJob) = "CP2" Or UCase(TheExec.CurrentJob) = "WLFT1" Then Exit Function
Dim SweepCondition_Split() As String: SweepCondition_Split = Split(argv(0), "+")
Dim sensor As String: sensor = argv(1)
Dim MetrologySense_ROT_Frequency As New DSPWave
Dim MetrologySense_ROV_Frequency As New DSPWave
Dim MTRSNS_Matrix_Index As Long: MTRSNS_Matrix_Index = argv(2)
Dim MTRSNS_Matrix_ROT_Row As Long: MTRSNS_Matrix_ROT_Row = argv(3)
Dim MTRSNS_Matrix_ROT_Column As Long: MTRSNS_Matrix_ROT_Column = argv(4)
Dim MTRSNS_Matrix_ROV_Row As Long: MTRSNS_Matrix_ROV_Row = argv(5)
Dim MTRSNS_Matrix_ROV_Column As Long: MTRSNS_Matrix_ROV_Column = argv(6)
Dim a1 As New DSPWave
Dim a2 As New DSPWave
Dim a1_Compression As New DSPWave
Dim a2_Compression As New DSPWave
Dim a1_Compression_eFuse As New DSPWave
Dim a2_Compression_eFuse As New DSPWave
Dim Array_a1_Compression_eFuse() As Double
Dim Array_a2_Compression_eFuse() As Double
Dim a1_Compression_eFuse_Store() As New DSPWave: ReDim a1_Compression_eFuse_Store(MTRSNS_Matrix_ROT_Row - 1)
Dim a2_Compression_eFuse_Store() As New DSPWave: ReDim a2_Compression_eFuse_Store(MTRSNS_Matrix_ROV_Row - 1)
Dim a1_LogicalCompare As New DSPWave
Dim a1_LogicalCompare_Array() As Double
Dim a2_LogicalCompare As New DSPWave
Dim a2_LogicalCompare_Array() As Double
Dim DSP_MetrologySense_ROT_Matrix As New DSPWave: DSP_MetrologySense_ROT_Matrix.data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROT_Matrix
Dim DSP_MetrologySense_ROV_Matrix As New DSPWave: DSP_MetrologySense_ROV_Matrix.data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROV_Matrix
Dim DSP_ROT_a_max_min As New DSPWave: DSP_ROT_a_max_min.data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROT_a_max_min_Matrix
Dim DSP_ROV_a_max_min As New DSPWave: DSP_ROV_a_max_min.data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROV_a_max_min_Matrix
Dim a1_max() As Double
Dim a1_min() As Double
Dim a2_max() As Double
Dim a2_min() As Double
Dim i As Long
Dim site As Variant
Dim Monotonicity_ROT_Check(0) As Long
Dim Monotonicity_ROV_Check(0) As Long
Dim Monotonicity_Delta As New DSPWave: Monotonicity_Delta.CreateConstant 20000000#, UBound(SweepCondition_Split), DspDouble
Dim Monotonicity_ROT_Check_DSP As New DSPWave
Dim Monotonicity_ROV_Check_DSP As New DSPWave
Dim TestNameInput As String
Dim DSPTempCal_1 As New DSPWave
Dim DSPTempCal_2 As New DSPWave
Dim DSPTempCal_3 As New DSPWave



For i = 0 To UBound(SweepCondition_Split)
    If i = 0 Then
        MetrologySense_ROT_Frequency = GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & sensor & "-sensor-ROT")
        MetrologySense_ROV_Frequency = GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & sensor & "-sensor-ROV")
    Else
        For Each site In TheExec.sites
            MetrologySense_ROT_Frequency = MetrologySense_ROT_Frequency.Concatenate(GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & sensor & "-sensor-ROT"))
            MetrologySense_ROV_Frequency = MetrologySense_ROV_Frequency.Concatenate(GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & sensor & "-sensor-ROV"))
        Next site
    End If
Next i
For Each site In TheExec.sites
'    Monotonicity_ROT_Check(0) = MetrologySense_ROT_Frequency.Select(0, , UBound(SweepCondition_Split)).LogicalCompare(LessThan, MetrologySense_ROT_Frequency.Select(1, , UBound(SweepCondition_Split))).CalcSum
'    Monotonicity_ROV_Check(0) = MetrologySense_ROV_Frequency.Select(0, , UBound(SweepCondition_Split)).LogicalCompare(LessThan, MetrologySense_ROV_Frequency.Select(1, , UBound(SweepCondition_Split))).CalcSum
'    Monotonicity_ROT_Check(0) = MetrologySense_ROT_Frequency.Select(1, , UBound(SweepCondition_Split)).Subtract(MetrologySense_ROT_Frequency.Select(0, , UBound(SweepCondition_Split))).LogicalCompare(GreaterThan, Monotonicity_Delta).CalcSum
'    Monotonicity_ROV_Check(0) = MetrologySense_ROV_Frequency.Select(1, , UBound(SweepCondition_Split)).Subtract(MetrologySense_ROV_Frequency.Select(0, , UBound(SweepCondition_Split))).LogicalCompare(GreaterThan, Monotonicity_Delta).CalcSum
'    If Monotonicity_ROT_Check(0) = UBound(SweepCondition_Split) Then Monotonicity_ROT_Check(0) = 0 Else Monotonicity_ROT_Check(0) = 1
'    If Monotonicity_ROV_Check(0) = UBound(SweepCondition_Split) Then Monotonicity_ROV_Check(0) = 0 Else Monotonicity_ROV_Check(0) = 1
'    Monotonicity_ROT_Check_DSP.Data = Monotonicity_ROT_Check
'    Monotonicity_ROV_Check_DSP.Data = Monotonicity_ROV_Check
    
    Monotonicity_ROT_Check_DSP = MetrologySense_ROT_Frequency.Select(1, , UBound(SweepCondition_Split)).Subtract(MetrologySense_ROT_Frequency.Select(0, , UBound(SweepCondition_Split)))
    Monotonicity_ROV_Check_DSP = MetrologySense_ROV_Frequency.Select(1, , UBound(SweepCondition_Split)).Subtract(MetrologySense_ROV_Frequency.Select(0, , UBound(SweepCondition_Split)))
Next site
For Each site In TheExec.sites
    MetrologySense_ROT_Frequency = MetrologySense_ROT_Frequency.divide(10 ^ 9)
    MetrologySense_ROV_Frequency = MetrologySense_ROV_Frequency.divide(10 ^ 9)
    a1 = DSP_MetrologySense_ROT_Matrix.MatrixMultiply(MTRSNS_Matrix_ROT_Row, MTRSNS_Matrix_ROT_Column, MetrologySense_ROT_Frequency)
    a2 = DSP_MetrologySense_ROV_Matrix.MatrixMultiply(MTRSNS_Matrix_ROV_Row, MTRSNS_Matrix_ROV_Column, MetrologySense_ROV_Frequency)
    a1_LogicalCompare_Array = a1.data
    a1_max = DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).data
    a1_min = DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row).data
    a2_LogicalCompare_Array = a2.data
    a2_max = DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).data
    a2_min = DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row).data
'    For i = 0 To UBound(a1_LogicalCompare_Array)
'        If a1_LogicalCompare_Array(i) > a1_max(i) Then
'            a1_LogicalCompare_Array(i) = a1_max(i)
'        ElseIf a1_LogicalCompare_Array(i) < a1_min(i) Then
'            a1_LogicalCompare_Array(i) = a1_min(i)
'        End If
'    Next i
'    For i = 0 To UBound(a2_LogicalCompare_Array)
'        If a2_LogicalCompare_Array(i) > a2_max(i) Then
'            a2_LogicalCompare_Array(i) = a2_max(i)
'        ElseIf a2_LogicalCompare_Array(i) < a2_min(i) Then
'            a2_LogicalCompare_Array(i) = a2_min(i)
'        End If
'    Next i
    a1_LogicalCompare.data = a1_LogicalCompare_Array
    a2_LogicalCompare.data = a2_LogicalCompare_Array
'''''''''''    a1_Compression = a1_LogicalCompare.Subtract(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row)).Divide(DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).Subtract(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row)))
'''''''''''    a2_Compression = a2_LogicalCompare.Subtract(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row)).Divide(DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).Subtract(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row)))
    
    
    
    DSPTempCal_1 = DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row).COPY
    DSPTempCal_2 = DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).COPY
    DSPTempCal_3 = DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row).COPY
    a1_Compression = a1_LogicalCompare.Subtract(DSPTempCal_1).divide(DSPTempCal_2.Subtract(DSPTempCal_3))
    
    DSPTempCal_1 = DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row).COPY
    DSPTempCal_2 = DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).COPY
    DSPTempCal_3 = DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row).COPY
    a2_Compression = a2_LogicalCompare.Subtract(DSPTempCal_1).divide(DSPTempCal_2.Subtract(DSPTempCal_3))
    
    
    
    
    Array_a1_Compression_eFuse = a1_Compression.data
    Array_a2_Compression_eFuse = a2_Compression.data
    For i = 0 To UBound(Array_a1_Compression_eFuse)
        If i = 0 Then
            If Array_a1_Compression_eFuse(i) >= 1 Then Array_a1_Compression_eFuse(i) = 2 ^ 15 - 1 Else Array_a1_Compression_eFuse(i) = FormatNumber(Array_a1_Compression_eFuse(i) * 2 ^ 15, 0)
        Else
            If Array_a1_Compression_eFuse(i) >= 1 Then Array_a1_Compression_eFuse(i) = 2 ^ 14 - 1 Else Array_a1_Compression_eFuse(i) = FormatNumber(Array_a1_Compression_eFuse(i) * 2 ^ 14, 0)
        End If
    Next i
    For i = 0 To UBound(Array_a2_Compression_eFuse)
        If i = 0 Then
            If Array_a2_Compression_eFuse(i) >= 1 Then Array_a2_Compression_eFuse(i) = 2 ^ 15 - 1 Else Array_a2_Compression_eFuse(i) = FormatNumber(Array_a2_Compression_eFuse(i) * 2 ^ 15, 0)
        Else
            If Array_a2_Compression_eFuse(i) >= 1 Then Array_a2_Compression_eFuse(i) = 2 ^ 14 - 1 Else Array_a2_Compression_eFuse(i) = FormatNumber(Array_a2_Compression_eFuse(i) * 2 ^ 14, 0)
        End If
    Next i
    a1_Compression_eFuse.data = Array_a1_Compression_eFuse
    a2_Compression_eFuse.data = Array_a2_Compression_eFuse
Next site
'For i = 0 To MTRSNS_Matrix_ROT_Column - 2
'    TestNameInput = Report_TName_From_Instance("CalcC", "", , 0)
'    TheExec.Flow.TestLimit resultVal:=Monotonicity_ROT_Check_DSP.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
'Next i
'For i = 0 To MTRSNS_Matrix_ROV_Column - 2
'    TestNameInput = Report_TName_From_Instance("CalcC", "", , 0)
'    TheExec.Flow.TestLimit resultVal:=Monotonicity_ROV_Check_DSP.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
'Next i
For i = 0 To MTRSNS_Matrix_ROT_Row - 1
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
    TheExec.flow.TestLimit resultVal:=a1.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
    TheExec.flow.TestLimit resultVal:=a1_Compression.Element(i), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Tname:=TestNameInput, ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        a1_Compression_eFuse_Store(i) = a1_Compression_eFuse.Select(i, , 1).COPY
    Next site
    Call StoreDataAllType("mtr_" & sensor & "_t1_a1_" & CStr(i + 1), a1_Compression_eFuse_Store(i))
Next i
For i = 0 To MTRSNS_Matrix_ROV_Row - 1
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
    TheExec.flow.TestLimit resultVal:=a2.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
    TheExec.flow.TestLimit resultVal:=a2_Compression.Element(i), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Tname:=TestNameInput, ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        a2_Compression_eFuse_Store(i) = a2_Compression_eFuse.Select(i, , 1).COPY
    Next site
    Call StoreDataAllType("mtr_" & sensor & "_t1_a2_" & CStr(i + 1), a2_Compression_eFuse_Store(i))
Next i

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologySense_Compression") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MetrologyTMPS_2s_Complement_Fractional_Conversion(ByRef Coeff_Dict As DSPWave, Coeff As DSPWave, Integer_Bit As Long, Fractional_Bit As Long) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim site As Variant
Dim Array_Coeff_Dict() As Double
Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Integer_Bit - 1, "1")) + Bin2Dec_rev_Fractional(String(Fractional_Bit, "1"))
Dim Low_limit As Double: Low_limit = -2 ^ (Integer_Bit - 1)

    For Each site In TheExec.sites
        Array_Coeff_Dict = Coeff.data
        If Array_Coeff_Dict(0) < Low_limit Then
            Array_Coeff_Dict(0) = 2 ^ (Integer_Bit + Fractional_Bit) + FormatNumber(2 ^ Fractional_Bit * Low_limit, 0)
        ElseIf Array_Coeff_Dict(0) >= Low_limit And Array_Coeff_Dict(0) < 0 Then
            If FormatNumber(2 ^ Fractional_Bit * Array_Coeff_Dict(0), 0) = 0 Then Array_Coeff_Dict(0) = 0 Else Array_Coeff_Dict(0) = 2 ^ (Integer_Bit + Fractional_Bit) + FormatNumber(2 ^ Fractional_Bit * Array_Coeff_Dict(0), 0)
        ElseIf Array_Coeff_Dict(0) < High_limit And Array_Coeff_Dict(0) >= 0 Then
            Array_Coeff_Dict(0) = FormatNumber(2 ^ Fractional_Bit * Array_Coeff_Dict(0), 0)
        Else
            Array_Coeff_Dict(0) = FormatNumber(2 ^ Fractional_Bit * High_limit, 0)
        End If
        Coeff_Dict.data = Array_Coeff_Dict
    Next site
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "MetrologyTMPS_2s_Complement_Fractional_Conversion") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MetrologySense_DeCompression(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    If UCase(TheExec.CurrentJob) = "CP1" Or UCase(TheExec.CurrentJob) = "CP2" Or UCase(TheExec.CurrentJob) = "WLFT1" Then Exit Function
    Dim SweepCondition_Split() As String: SweepCondition_Split = Split(argv(0), "+")
    Dim sensor As String: sensor = argv(1)
    Dim MTRSNS_Matrix_Index As Long: MTRSNS_Matrix_Index = argv(2)
    Dim MTRSNS_Matrix_ROT_Row As Long: MTRSNS_Matrix_ROT_Row = argv(3)
    Dim MTRSNS_Matrix_ROT_Column As Long: MTRSNS_Matrix_ROT_Column = argv(4)
    Dim MTRSNS_Matrix_ROV_Row As Long: MTRSNS_Matrix_ROV_Row = argv(5)
    Dim MTRSNS_Matrix_ROV_Column As Long: MTRSNS_Matrix_ROV_Column = argv(6)
    Dim MetrologySense_ROT_Frequency As New DSPWave
    Dim MetrologySense_ROV_Frequency As New DSPWave
    Dim MetrologySense_ROT_Frequency_DeCompression As New DSPWave
    Dim MetrologySense_ROV_Frequency_DeCompression As New DSPWave
'    Dim MetrologySense_ROT_Frequency_DeCompression_Store() As New DSPWave: ReDim MetrologySense_ROT_Frequency_DeCompression_Store(UBound(SweepCondition_Split))
'    Dim MetrologySense_ROV_Frequency_DeCompression_Store() As New DSPWave: ReDim MetrologySense_ROV_Frequency_DeCompression_Store(UBound(SweepCondition_Split))
    Dim DSP_MetrologySense_ROT_Matrix As New DSPWave: DSP_MetrologySense_ROT_Matrix.data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROT_Matrix
    Dim DSP_MetrologySense_ROV_Matrix As New DSPWave: DSP_MetrologySense_ROV_Matrix.data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROV_Matrix
    Dim a1_Compression_eFuse As New DSPWave
    Dim a2_Compression_eFuse As New DSPWave
    Dim a1_Compression As New DSPWave
    Dim a2_Compression As New DSPWave
    Dim Array_a1_Compression() As Double
    Dim Array_a2_Compression() As Double
    Dim Array_a1_Compression_Str() As String: ReDim Array_a1_Compression_Str(MTRSNS_Matrix_ROT_Row - 1)
    Dim Array_a2_Compression_Str() As String: ReDim Array_a2_Compression_Str(MTRSNS_Matrix_ROV_Row - 1)
    Dim a1_DeCompression As New DSPWave
    Dim a2_DeCompression As New DSPWave
    Dim DSP_ROT_a_max_min As New DSPWave: DSP_ROT_a_max_min.data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROT_a_max_min_Matrix
    Dim DSP_ROV_a_max_min As New DSPWave: DSP_ROV_a_max_min.data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROV_a_max_min_Matrix
    Dim MetrologySense_ROT_Frequency_Error As New DSPWave
    Dim MetrologySense_ROV_Frequency_Error As New DSPWave
'    Dim FlowLimitObj As IFlowLimitsInfo: Call TheExec.Flow.GetTestLimits(FlowLimitObj)
'    Dim FlowTestName() As String: Call FlowLimitObj.GetTNames(FlowTestName)
'    Dim TestLimitIndex As Long: TestLimitIndex = TheExec.Flow.TestLimitIndex
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    For i = 0 To UBound(SweepCondition_Split)
        If i = 0 Then
            MetrologySense_ROT_Frequency = GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & sensor & "-sensor-ROT")
            MetrologySense_ROV_Frequency = GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & sensor & "-sensor-ROV")
        Else
            For Each site In TheExec.sites
                MetrologySense_ROT_Frequency = MetrologySense_ROT_Frequency.Concatenate(GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & sensor & "-sensor-ROT"))
                MetrologySense_ROV_Frequency = MetrologySense_ROV_Frequency.Concatenate(GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & sensor & "-sensor-ROV"))
            Next site
        End If
    Next i
    For i = 0 To MTRSNS_Matrix_ROT_Row - 1
        If i = 0 Then
            a1_Compression_eFuse = GetStoreDataAllType("mtr_" & sensor & "_t1_a1_" & CStr(i + 1))
        Else
            For Each site In TheExec.sites
                a1_Compression_eFuse = a1_Compression_eFuse.Concatenate(GetStoreDataAllType("mtr_" & sensor & "_t1_a1_" & CStr(i + 1)))
            Next site
        End If
    Next i
    For i = 0 To MTRSNS_Matrix_ROV_Row - 1
        If i = 0 Then
            a2_Compression_eFuse = GetStoreDataAllType("mtr_" & sensor & "_t1_a2_" & CStr(i + 1))
        Else
            For Each site In TheExec.sites
                a2_Compression_eFuse = a2_Compression_eFuse.Concatenate(GetStoreDataAllType("mtr_" & sensor & "_t1_a2_" & CStr(i + 1)))
            Next site
        End If
    Next i
    For Each site In TheExec.sites
        Array_a1_Compression = a1_Compression_eFuse.data
        Array_a2_Compression = a2_Compression_eFuse.data
        For i = 0 To UBound(Array_a1_Compression)
            If i = 0 Then
                Call Dec2Bin_str(Array_a1_Compression(i), Array_a1_Compression_Str(i), 14)
            Else
                Call Dec2Bin_str(Array_a1_Compression(i), Array_a1_Compression_Str(i), 13)
            End If
            Array_a1_Compression(i) = Bin2Dec_rev_Fractional(Array_a1_Compression_Str(i))
        Next i
        For i = 0 To UBound(Array_a2_Compression)
            If i = 0 Then
                Call Dec2Bin_str(Array_a2_Compression(i), Array_a2_Compression_Str(i), 14)
            Else
                Call Dec2Bin_str(Array_a2_Compression(i), Array_a2_Compression_Str(i), 13)
            End If
            Array_a2_Compression(i) = Bin2Dec_rev_Fractional(Array_a2_Compression_Str(i))
        Next i
        a1_Compression.data = Array_a1_Compression
        a2_Compression.data = Array_a2_Compression
        a1_DeCompression = a1_Compression.Multiply(DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).Subtract(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row))).Add(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row).COPY)
        a2_DeCompression = a2_Compression.Multiply(DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).Subtract(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row))).Add(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row).COPY)
        MetrologySense_ROT_Frequency_DeCompression = DSP_MetrologySense_ROT_Matrix.MatrixTranspose(MTRSNS_Matrix_ROT_Row).MatrixMultiply(MTRSNS_Matrix_ROT_Column, MTRSNS_Matrix_ROT_Row, a1_DeCompression)
        MetrologySense_ROV_Frequency_DeCompression = DSP_MetrologySense_ROV_Matrix.MatrixTranspose(MTRSNS_Matrix_ROV_Row).MatrixMultiply(MTRSNS_Matrix_ROV_Column, MTRSNS_Matrix_ROV_Row, a2_DeCompression)
        MetrologySense_ROT_Frequency_DeCompression = MetrologySense_ROT_Frequency_DeCompression.Multiply(10 ^ 9)
        MetrologySense_ROV_Frequency_DeCompression = MetrologySense_ROV_Frequency_DeCompression.Multiply(10 ^ 9)
        MetrologySense_ROT_Frequency_Error = MetrologySense_ROT_Frequency_DeCompression.Subtract(MetrologySense_ROT_Frequency).divide(MetrologySense_ROT_Frequency_DeCompression)
        MetrologySense_ROV_Frequency_Error = MetrologySense_ROV_Frequency_DeCompression.Subtract(MetrologySense_ROV_Frequency).divide(MetrologySense_ROV_Frequency_DeCompression)
    Next site
'    If UCase(TheExec.CurrentJob) = "FT2" Then
'        For i = 0 To UBound(SweepCondition_Split)
'            For Each site In TheExec.sites
'                MetrologySense_ROT_Frequency_DeCompression_Store(i) = MetrologySense_ROT_Frequency_DeCompression.Select(i, , 1)
'                MetrologySense_ROV_Frequency_DeCompression_Store(i) = MetrologySense_ROV_Frequency_DeCompression.Select(i, , 1)
'            Next site
'            Call StoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT-DeCompression", MetrologySense_ROT_Frequency_DeCompression_Store(i))
'            Call StoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV-DeCompression", MetrologySense_ROV_Frequency_DeCompression_Store(i))
'        Next i
'    End If
'    Do While Not (LCase(FlowTestName(TestLimitIndex)) Like "*decompression*")
'        TestLimitIndex = TestLimitIndex + 1
'    Loop
'    TheExec.Flow.TestLimitIndex = TestLimitIndex
    For i = 0 To MTRSNS_Matrix_ROT_Column - 1
        TestNameInput = Report_TName_From_Instance("CalcF", vbNullString, , CInt(i))
        TheExec.flow.TestLimit resultVal:=MetrologySense_ROT_Frequency_DeCompression.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcF", vbNullString, , CInt(i))
        TheExec.flow.TestLimit resultVal:=MetrologySense_ROV_Frequency_DeCompression.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
        TheExec.flow.TestLimit resultVal:=MetrologySense_ROT_Frequency_Error.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
        TheExec.flow.TestLimit resultVal:=MetrologySense_ROV_Frequency_Error.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologySense_DeCompression") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyGR_Offset(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim MTRGR_Offset As New SiteDouble: MTRGR_Offset = GetStoreDataAllType(argv(0) & "_para")
    Dim Integer_Bit As Long: Integer_Bit = argv(1)
    Dim Dictionary_Name As String: Dictionary_Name = argv(2)
    Dim MTRGR_Offset_Array(0) As Double
    Dim DSP_MTRGR_Offset As New DSPWave
    Dim DSP_MTRGR_Offset_eFuse As New DSPWave
    
    For Each site In TheExec.sites
        MTRGR_Offset_Array(0) = MTRGR_Offset
        DSP_MTRGR_Offset.data = MTRGR_Offset_Array
    Next site
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_MTRGR_Offset_eFuse, DSP_MTRGR_Offset, Integer_Bit, 0)
    Call StoreDataAllType(Dictionary_Name, DSP_MTRGR_Offset_eFuse)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyGR_Offset") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Function Calc_MetrologyGR_Gain(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim MTRGR_Gain_Split() As String: MTRGR_Gain_Split = Split(argv(0), "+")
    Dim MTRGR_Gain() As New SiteDouble: ReDim MTRGR_Gain(UBound(MTRGR_Gain_Split))
    Dim MTRGR_Offset_Dictionary As String: MTRGR_Offset_Dictionary = argv(1) & "_para"
    Dim DictionaryName As String: DictionaryName = argv(2)
    Dim MTRGR_Offset As New SiteDouble: MTRGR_Offset = GetStoreDataAllType(MTRGR_Offset_Dictionary)
    Dim DSP_MTRGR_Gain_eFuse As New DSPWave
    Dim MTRGR_Gain_eFuse_Array(0) As Double
    Dim i As Long
    Dim TestNameInput As String
    
    For i = 0 To UBound(MTRGR_Gain_Split)
        MTRGR_Gain(i) = GetStoreDataAllType(MTRGR_Gain_Split(i) & "_para")
    Next i
    For Each site In TheExec.sites
        MTRGR_Gain_eFuse_Array(0) = 0
        For i = 0 To UBound(MTRGR_Gain_Split)
            MTRGR_Gain_eFuse_Array(0) = MTRGR_Gain_eFuse_Array(0) + MTRGR_Gain(i)
        Next i
        MTRGR_Gain_eFuse_Array(0) = MTRGR_Gain_eFuse_Array(0) - 8 * MTRGR_Offset
        DSP_MTRGR_Gain_eFuse.data = MTRGR_Gain_eFuse_Array
    Next site

    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_MTRGR_Gain_eFuse.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Call StoreDataAllType(DictionaryName, DSP_MTRGR_Gain_eFuse)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyGR_Gain") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyTMPS_OffSet_Update(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
Dim DSP_OffSet_Mean As New DSPWave
Dim DSP_OffSet_Mean_Array(0) As Double
Dim DSP_OffSet_Mean_eFuse As New DSPWave
'Dim DSP_OffSet_Mean_Fuse_Array(0) As Double
Dim TestNameInput As String
Dim i As Long
    For i = 0 To UBound(InWf_Array)
        InWf_SiteDouble(i) = GetStoreDataAllType(InWf_Split(i) & "_para")
    Next i
    For Each site In TheExec.sites.Active
        For i = 0 To UBound(InWf_Array)
            InWf_Array(i) = InWf_SiteDouble(i)
        Next i
        InWf.data = InWf_Array
        DSP_OffSet_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0)
        DSP_OffSet_Mean.data = DSP_OffSet_Mean_Array
    Next site

    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=DSP_OffSet_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
'    Call StoreDataAllType(argv(1) & "_" & CStr(TheExec.Flow.var(argv(2)).Value), DSP_OffSet_Mean)
    Call StoreDataAllType(argv(1) & "_15", DSP_OffSet_Mean)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_OffSet_Mean_eFuse, DSP_OffSet_Mean, 18, 0)
'    Call StoreDataAllType(argv(1) & "_eFuse_" & CStr(TheExec.Flow.var(argv(2)).Value), DSP_OffSet_Mean_eFuse)
    Call StoreDataAllType(argv(1) & "_eFuse_15", DSP_OffSet_Mean_eFuse)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyTMPS_OffSet_Update") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyTMPS_Gain_Update(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")

Dim DSP_Gain_Mean As New DSPWave
Dim DSP_Gain_Mean_Array(0) As Double
Dim DSP_OffSet_Mean As New DSPWave
Dim DSP_OffSet_Mean_Array() As Double
Dim DSP_Gain_Mean_Final As New DSPWave
Dim DSP_Gain_Mean_Final_Array(0) As Double
Dim TestNameInput As String
Dim i As Long

For i = 0 To UBound(InWf_Array)
    InWf_SiteDouble(i) = GetStoreDataAllType(InWf_Split(i) & "_para")
Next i

'DSP_OffSet_Mean = GetStoreDataAllType(argv(1) & "_" & CStr(TheExec.Flow.var(argv(2)).Value))
DSP_OffSet_Mean = GetStoreDataAllType(argv(1) & "_15")
For Each site In TheExec.sites.Active
    DSP_OffSet_Mean_Array = DSP_OffSet_Mean.data
    For i = 0 To UBound(InWf_Array)
        InWf_Array(i) = InWf_SiteDouble(i)
    Next i
    InWf.data = InWf_Array
    DSP_Gain_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0) - DSP_OffSet_Mean_Array(0)
    If DSP_Gain_Mean_Array(0) < 0 Then: DSP_Gain_Mean_Array(0) = 0
    DSP_Gain_Mean.data = DSP_Gain_Mean_Array
Next site

TestNameInput = Report_TName_From_Instance("CalcC", "X", Replace(argv(3), "_", vbNullString), 0, 0)
'TestNameInput = Report_TName_From_Instance("CalcC", "", Replace(argv(3) & "_AVG", "_", ""))
TheExec.flow.TestLimit resultVal:=DSP_Gain_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

Call StoreDataAllType(argv(3) & "_" & CStr(TheExec.flow.var(argv(2)).value), DSP_Gain_Mean)

If TheExec.flow.var(argv(2)).value = argv(5) Then
    For Each site In TheExec.sites.Active
        InWf.Clear
        For i = argv(4) To argv(5)
            If i = argv(4) Then
                InWf = GetStoreDataAllType(argv(3) & "_" & i)
            Else
                InWf = InWf.Concatenate(GetStoreDataAllType(argv(3) & "_" & i))
            End If
        Next i

    DSP_Gain_Mean_Final_Array(0) = FormatNumber(InWf.CalcMean, 0)
    DSP_Gain_Mean_Final.data = DSP_Gain_Mean_Final_Array
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_AVG", "_", vbNullString))
    TheExec.flow.TestLimit resultVal:=DSP_Gain_Mean_Final.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    Call StoreDataAllType(argv(3), DSP_Gain_Mean_Final)
End If

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyTMPS_Gain_Update") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MetrologyBTS_OffSet(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim Sensor_Num() As String
Dim TestNameInput As String
Dim BTS_V0 As New SiteDouble
Dim i As Long
    Sensor_Num = Split(argv(1), "_")
    
    InWf_SiteDouble(0) = GetStoreDataAllType(argv(0) & "_para")
   ' BTS_V0 = InWf_SiteDouble(0).Divide(2 ^ 16)
    BTS_V0 = InWf_SiteDouble(0).Subtract(2 ^ 16).divide(2 ^ 16) '.Abs.Multiply(0) 'V05A pipe
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=BTS_V0, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling ', formatStr:="%.5f"   ''new
       
    Call StoreDataAllType(argv(1), BTS_V0)  '& "_" & CStr(TheExec.Flow.var(argv(2)).Value)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_OffSet") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyBTS_OffSetB0(argc As Integer, argv() As String) As Long
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim Sensor_Num() As String
Dim TestNameInput As String
Dim BTS_V0 As New SiteDouble
Dim i As Long
    Sensor_Num = Split(argv(1), "_")
    
    InWf_SiteDouble(0) = GetStoredData(argv(0) & "_para")
   ' BTS_V0 = InWf_SiteDouble(0).Divide(2 ^ 16)
    BTS_V0 = InWf_SiteDouble(0).Subtract(2 ^ 14).divide(2 ^ 14) '.Abs.Multiply(0) 'V05A pipe
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=BTS_V0, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling ', formatStr:="%.5f"   ''new
       
    Call AddStoredData(argv(1), BTS_V0)  '& "_" & CStr(TheExec.Flow.var(argv(2)).Value)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_OffSet"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Calc_MetrologyBTS_OffSetB0_Ibiza(argc As Integer, argv() As String) As Long
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim Sensor_Num() As String
Dim TestNameInput As String
Dim BTS_V0 As New SiteDouble
Dim i As Long
    Sensor_Num = Split(argv(1), "_")
    
    InWf_SiteDouble(0) = GetStoredData(argv(0) & "_para")
    BTS_V0 = InWf_SiteDouble(0).divide(2 ^ 16)
    'BTS_V0 = InWf_SiteDouble(0).Subtract(2 ^ 14).divide(2 ^ 14) '.Abs.Multiply(0) 'V05A pipe
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=BTS_V0, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling ', formatStr:="%.5f"   ''new
       
    Call AddStoredData(argv(1), BTS_V0)  '& "_" & CStr(TheExec.Flow.var(argv(2)).Value)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_OffSet"
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Calc_MetrologyBTS_Gain(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim DSP_MTRBTS_OUT1 As New SiteDouble: DSP_MTRBTS_OUT1 = GetStoreDataAllType(argv(0) & "_para")
    Dim DSP_MTRBTS_OUT2 As New SiteDouble: DSP_MTRBTS_OUT2 = GetStoreDataAllType(argv(1) & "_para")

    Dim DSP_MTRBTS_OUT1_DEC As New DSPWave
    Dim DSP_MTRBTS_OUT2_DEC As New DSPWave
    Dim DSP_MTRBTS_GAIN_DEC As New DSPWave
    Dim DSP_MTRBTS_GAIN_BINARY As New DSPWave
    Dim ARRAY_MTRBTS_OUT1_DEC(0) As Double
    Dim ARRAY_MTRBTS_OUT2_DEC(0) As Double
    Dim TestNameInput As String
    Dim BTS_V1 As New SiteDouble
    Dim Sensor_Num() As String
    Sensor_Num = Split(argv(2), "_")
    Dim i As Integer
    For Each site In TheExec.sites.Active
        ARRAY_MTRBTS_OUT1_DEC(0) = DSP_MTRBTS_OUT1
        ARRAY_MTRBTS_OUT2_DEC(0) = DSP_MTRBTS_OUT2
        
        DSP_MTRBTS_OUT1_DEC.data = ARRAY_MTRBTS_OUT1_DEC
        DSP_MTRBTS_OUT2_DEC.data = ARRAY_MTRBTS_OUT2_DEC
        BTS_V1 = DSP_MTRBTS_OUT1_DEC.Subtract(DSP_MTRBTS_OUT2_DEC).divide(2 ^ 16).Element(0)
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=BTS_V1, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.4f" ''try
    Call StoreDataAllType(argv(2), BTS_V1)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Gain") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_MetrologyBTS_Coefficient(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim DSP_MTRBTS_OUT As New DSPWave
    Dim ARRAY_MTRBTS_OUT_DEC() As New SiteDouble
    Dim v0 As New SiteDouble
    Dim v1 As New SiteDouble
    ReDim ARRAY_MTRBTS_OUT_DEC(0)
    Dim x0(0) As New SiteDouble
    Dim x0a(0) As New SiteDouble
    Dim GC0, GC1, GC2, GC3 As Double
    Dim C0_CAL(0) As New SiteDouble
    Dim A0_CAL(0) As New SiteDouble
    Dim B0_CAL(0) As New SiteDouble
    Dim B1_CAL(0) As New SiteDouble
    Dim ARRAY_C0_CAL_eFuse() As Double
    Dim ARRAY_C0_CAL_SRC() As Long: ReDim ARRAY_C0_CAL_SRC(24)
    Dim TempVal As Double
    Dim TestNameInput As String
    Dim i As Long
    Dim Sensor_Num() As String
    Sensor_Num = Split(argv(0), "_")
    
    '===================201209 BTS new constant===========================
    Dim a0, a1, a2, a3, Tcp, k0, k1 As Double
        
    'V1
'    A0 = 53.42539937
'    a1 = 106.24914282
'    a2 = -5.64960584
'    a3 = 0.5126488
'    k0 = 133.35292861
'    k1 = 98.97608898
'    Tcp = 24.3
    
'    'V2 High Accuracy mode
    a0 = 53.26973036
    a1 = 104.64623585
    a2 = -8.85459385
    a3 = -0.30123576
    k0 = 131.34112566
    k1 = 97.48290549
    Tcp = 25

    If gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment "a0=" & a0 & ",a1=" & a1 & ",a2=" & a2 & ",a3=" & a3 & ",k0=" & k0 & ",k1=" & k1 & ",Tcp=" & Tcp
    End If
    ''===================================================================
    ARRAY_MTRBTS_OUT_DEC(0) = GetStoreDataAllType(argv(0) & "_para")
    v0 = GetStoreDataAllType(argv(1))
    v1 = GetStoreDataAllType(argv(2))
    For Each site In TheExec.sites.Active
        If v1 = 0 Then
            v1 = 0.00000001
            TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Coefficient V1 value is zero!!! "
        End If
        x0a(0) = ARRAY_MTRBTS_OUT_DEC(0) / (2 ^ 16) ''201209 for new BTS rule
        x0(0) = (x0a(0) - v0) / v1
        A0_CAL(0) = Tcp - a1 * x0(0) - a2 * x0(0) ^ 2 - a3 * x0(0) ^ 3
        B1_CAL(0) = k0 + k1 * x0(0)
        B0_CAL(0) = Tcp - B1_CAL(0) * x0(0) - a2 * x0(0) ^ 2 - a3 * x0(0) ^ 3
    Next site
    '===================201209 BTS new rule===========================
    Dim C0 As New SiteDouble
    Dim C1 As New SiteDouble
    Dim C2 As New SiteDouble
    Dim C3 As New SiteDouble
    Dim C0a As New SiteDouble
    Dim C1a As New SiteDouble
    Dim DSP_C0_Cal_eFuse As New DSPWave
    Dim DSP_C1_Cal_eFuse As New DSPWave
    Dim DSP_C2_Cal_eFuse As New DSPWave
    Dim DSP_C3_Cal_eFuse As New DSPWave
    Dim DSP_C0a_Cal_eFuse As New DSPWave
    Dim DSP_C1a_Cal_eFuse As New DSPWave
    Dim DSP_C0_CAL  As New DSPWave
    Dim DSP_C1_CAL  As New DSPWave
    Dim DSP_C2_CAL  As New DSPWave
    Dim DSP_C3_CAL  As New DSPWave
    Dim DSP_C0a_CAL As New DSPWave
    Dim DSP_C1a_CAL As New DSPWave
    Dim DSP_C0_Cal_Src As New DSPWave
    Dim DSP_C1_Cal_Src As New DSPWave
    Dim DSP_C2_Cal_Src As New DSPWave
    Dim DSP_C3_Cal_Src As New DSPWave
    Dim DSP_C0a_Cal_Src As New DSPWave
    Dim DSP_C1a_Cal_Src As New DSPWave
    
    DSP_C0_Cal_eFuse.CreateConstant 0, 1
    DSP_C1_Cal_eFuse.CreateConstant 0, 1
    DSP_C2_Cal_eFuse.CreateConstant 0, 1
    DSP_C3_Cal_eFuse.CreateConstant 0, 1
    DSP_C0a_Cal_eFuse.CreateConstant 0, 1
    DSP_C1a_Cal_eFuse.CreateConstant 0, 1
    DSP_C0_CAL.CreateConstant 0, 1
    DSP_C1_CAL.CreateConstant 0, 1
    DSP_C2_CAL.CreateConstant 0, 1
    DSP_C3_CAL.CreateConstant 0, 1
    DSP_C0a_CAL.CreateConstant 0, 1
    DSP_C1a_CAL.CreateConstant 0, 1

    
    C0 = A0_CAL(0).Subtract(v0.Multiply(a1).divide(v1))                                    ''A0_CAL(0)- a1 * V0 / V1
    C1 = v1.Invert.Multiply(a1).Subtract(v0.Multiply(2 * a2).divide(v1.power(2)))          ''a1 / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
    C2 = v1.power(2).Invert.Multiply(a2).Subtract(v0.Multiply(3 * a3).divide(v1.power(3))) ''a2 / (BTS_V1 ^ 2) - 3 * a3 * BTS_V0 / (BTS_V1 ^ 3)
    C3 = v1.power(3).Invert.Multiply(a3)                                                   ''a3 / (BTS_V1 ^ 3)
    C0a = B0_CAL(0).Subtract(B1_CAL(0).Multiply(v0).divide(v1))                            ''B0_CAL(0) - B1_CAL(0) * BTS_V0 / BTS_V1
    C1a = B1_CAL(0).divide(v1).Subtract(v0.Multiply(2 * a2).divide(v1.power(2)))           ''B1_CAL / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
    '=================================================================
    For Each site In TheExec.sites.Active
        DSP_C0_CAL.Element(0) = C0
        DSP_C1_CAL.Element(0) = C1
        DSP_C2_CAL.Element(0) = C2
        DSP_C3_CAL.Element(0) = C3
        DSP_C0a_CAL.Element(0) = C0a
        DSP_C1a_CAL.Element(0) = C1a
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=x0a(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=x0(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=A0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=B0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=B1_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C0_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C0_SRC", DSP_C0_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C1_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C1_SRC", DSP_C1_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C2_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C2_SRC", DSP_C2_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C3_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C3_SRC", DSP_C3_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C0a_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C0A_SRC", DSP_C0A_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C1a_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C1A_SRC", DSP_C1A_Cal_Src)
    
    C0 = C0.Multiply(2 ^ 9)
    C1 = C1.Multiply(2 ^ 9)
    C2 = C2.Multiply(2 ^ 9)
    C3 = C3.Multiply(2 ^ 9)
    C0a = C0a.Multiply(2 ^ 9)
    C1a = C1a.Multiply(2 ^ 9)
    For Each site In TheExec.sites.Active
        DSP_C0_CAL.Element(0) = C0
        DSP_C1_CAL.Element(0) = C1
        DSP_C2_CAL.Element(0) = C2
        DSP_C3_CAL.Element(0) = C3
        DSP_C0a_CAL.Element(0) = C0a
        DSP_C1a_CAL.Element(0) = C1a
    Next site
    
''    For Each site In TheExec.sites.Active
''        DSP_C0_CAL.Element(0) = 0 'C0
''        DSP_C1_CAL.Element(0) = 0 'C0 '1
''        DSP_C2_CAL.Element(0) = 0 'C0 '2
''        DSP_C3_CAL.Element(0) = 0 'C3
''        DSP_C0A_CAL.Element(0) = 0 ' C0A
''        DSP_C1A_CAL.Element(0) = 0 ' C1A
''    Next site
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0_Cal_eFuse, DSP_C0_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1_Cal_eFuse, DSP_C1_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C2_Cal_eFuse, DSP_C2_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C3_Cal_eFuse, DSP_C3_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0a_Cal_eFuse, DSP_C0a_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1a_Cal_eFuse, DSP_C1a_CAL, 20, 0)
    
    Call HardIP_Dec2Bin(DSP_C0_Cal_Src, DSP_C0_Cal_eFuse, 20) 'DSP_DecToBin
    Call HardIP_Dec2Bin(DSP_C1_Cal_Src, DSP_C1_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C2_Cal_Src, DSP_C2_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C3_Cal_Src, DSP_C3_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C0a_Cal_Src, DSP_C0a_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C1a_Cal_Src, DSP_C1a_Cal_eFuse, 20)
    
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites.Active
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C0=" & DSP_C0_CAL.Element(0) & " => " & DSP_C0_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C1=" & DSP_C1_CAL.Element(0) & " => " & DSP_C1_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C2=" & DSP_C2_CAL.Element(0) & " => " & DSP_C2_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C3=" & DSP_C3_CAL.Element(0) & " => " & DSP_C3_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C0A=" & DSP_C0a_CAL.Element(0) & " => " & DSP_C0a_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C1A=" & DSP_C1a_CAL.Element(0) & " => " & DSP_C1a_Cal_eFuse.Element(0)
        Next site
    End If
    
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=x0a(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=x0(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=A0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=B0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=B1_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=DSP_C0_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    Call StoreDataAllType(Sensor_Num(0) & "_C0_SRC", DSP_C0_Cal_Src)
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=DSP_C1_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    Call StoreDataAllType(Sensor_Num(0) & "_C1_SRC", DSP_C1_Cal_Src)
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=DSP_C2_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    Call StoreDataAllType(Sensor_Num(0) & "_C2_SRC", DSP_C2_Cal_Src)
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=DSP_C3_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    Call StoreDataAllType(Sensor_Num(0) & "_C3_SRC", DSP_C3_Cal_Src)
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=DSP_C0A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    Call StoreDataAllType(Sensor_Num(0) & "_C0A_SRC", DSP_C0a_Cal_Src)
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=DSP_C1A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    Call StoreDataAllType(Sensor_Num(0) & "_C1A_SRC", DSP_C1a_Cal_Src)
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Coefficient") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MetrologyBTS_Temperature(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    Dim Sensor_Num() As String
    
    Dim DSP_MTRSNS_Temperature_Average As New DSPWave
    Dim DSP_MTRSNS_Temperature_Maximum As New DSPWave
    Dim DSP_MTRSNS_Temperature_Minimum As New DSPWave
    Dim DSP_MTRSNS_Temperature_Maximum_2 As New DSPWave
    
    DSP_MTRSNS_Temperature_Average.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Maximum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Minimum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Maximum_2.CreateConstant 0, UBound(Temperature_Sensor) + 1
    
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
        If argc > 1 Then
            Call AddStoredData(Temperature_Sensor(i) & "_" & argv(1), Temperature.divide(64))
        End If
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).data = Temperature_Array

            If UCase(TheExec.flow.CurrentFlowSheetName) Like "FLOW_TMPS*" Or UCase(TheExec.flow.CurrentFlowSheetName) = "FLOW_TMPS_NO_RELAY" Or UCase(TheExec.flow.CurrentFlowSheetName) = "FLOW_TMPS_TSNS" Or UCase(TheExec.flow.CurrentFlowSheetName) Like "FLOW_TMPS_EVS*" Then
                DSP_MTRSNS_Temperature_Average.Element(i) = DSP_MTRSNS_Temperature_Average.Element(i) + DSP_Temperature(i).Element(0)
                DSP_MTRSNS_Temperature_Maximum.Element(i) = DSP_MTRSNS_Temperature_Maximum.Element(i) + DSP_Temperature(i).Element(0)
                DSP_MTRSNS_Temperature_Minimum.Element(i) = DSP_MTRSNS_Temperature_Minimum.Element(i) + DSP_Temperature(i).Element(0)
                DSP_MTRSNS_Temperature_Maximum_2.Element(i) = DSP_MTRSNS_Temperature_Maximum_2.Element(i) + DSP_Temperature(i).Element(0)
            End If
        Next site
    Next i

    
    If UCase(TheExec.flow.CurrentFlowSheetName) Like "FLOW_TMPS*" Or UCase(TheExec.flow.CurrentFlowSheetName) = "FLOW_TMPS_NO_RELAY" Or UCase(TheExec.flow.CurrentFlowSheetName) = "FLOW_TMPS_TSNS" Or UCase(TheExec.flow.CurrentFlowSheetName) Like "FLOW_TMPS_EVS*" Then
    Else
        For i = 0 To UBound(Temperature_Sensor)
            Sensor_Num = Split(Temperature_Sensor(i), "_")
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow

        Next i
    End If
    
    If UCase(TheExec.flow.CurrentFlowSheetName) Like "FLOW_TMPS*" Or UCase(TheExec.flow.CurrentFlowSheetName) = "FLOW_TMPS_NO_RELAY" Or UCase(TheExec.flow.CurrentFlowSheetName) = "FLOW_TMPS_TSNS" Or UCase(TheExec.flow.CurrentFlowSheetName) Like "FLOW_TMPS_EVS*" Then
        For Each site In TheExec.sites
            DSP_MTRSNS_Temperature_Average(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Average(site).CalcMean, 4)
            DSP_MTRSNS_Temperature_Maximum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Maximum(site).CalcMaximumValue, 4)
            DSP_MTRSNS_Temperature_Minimum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Minimum(site).CalcMinimumValue, 4)
            DSP_MTRSNS_Temperature_Maximum_2(site).Element(0) = DSP_MTRSNS_Temperature_Maximum_2(site).CalcMaximumValue
        Next site
        
        TMPS_Fail = False
        For Each site In TheExec.sites
            If Instance_Data.hiLimit(TheExec.flow.TestLimitIndex) = "" Then
            ElseIf DSP_MTRSNS_Temperature_Maximum_2.Element(0) > CDbl(Instance_Data.hiLimit(TheExec.flow.TestLimitIndex)) Then
                TMPS_Fail = True
                Exit For
            End If
        Next site
    
        If TheExec.enableWord("Enable_Vddbinning_Cooling") = False Then
          glb_TMPS_count_str = "Last"
          glb_TMPS_count_int = 0
          Time0 = Timer
        Else
            If TheExec.flow.var("SrcCodeIndxZ").value = 1 Then
                glb_TMPS_count_str = "First"
                glb_TMPS_count_int = 0
                Time0 = Timer
            ElseIf TMPS_Fail = False Then
                glb_TMPS_count_str = "Last"
                glb_TMPS_count_int = glb_TMPS_count_int + 1
            ElseIf TheExec.flow.var("SrcCodeIndxZ").value = 7 Then
                glb_TMPS_count_str = "Last"
                glb_TMPS_count_int = glb_TMPS_count_int + 1
            Else
                glb_TMPS_count_int = glb_TMPS_count_int + 1
                glb_TMPS_count_str = CStr(glb_TMPS_count_int)
            End If
        End If
    
    
        For i = 0 To UBound(Temperature_Sensor)
            Sensor_Num = Split(Temperature_Sensor(i), "_")
            TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
            'TheExec.Flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_" & Replace(Split(TestNameInput, "_")(6), "-", "_") & "_" & glb_TMPS_count_str, ForceResults:=tlForceFlow, customUnit:="C"
        Next i
    
        TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature_Average" & "_" & glb_TMPS_count_str, ForceResults:=tlForceNone, customUnit:="C"
        TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature_Maximum" & "_" & glb_TMPS_count_str, ForceResults:=tlForceNone, customUnit:="C"
        TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature_Minimum" & "_" & glb_TMPS_count_str, ForceResults:=tlForceNone, customUnit:="C"
        
        
        If TheExec.enableWord("Enable_Vddbinning_Cooling") = True Then
        
            If TheExec.flow.var("SrcCodeIndxZ").value = 1 And TMPS_Fail = False Then
                TheExec.flow.TestLimit resultVal:=0, unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature" & "_" & "CoolingLoopCount", ForceResults:=tlForceNone
                TheExec.flow.TestLimit resultVal:=0, unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature" & "_" & "CoolingTime", ForceResults:=tlForceNone, customUnit:="sec"
            ElseIf (TheExec.flow.var("SrcCodeIndxZ").value <> 1 And TMPS_Fail = False) Or TheExec.flow.var("SrcCodeIndxZ").value = 7 Then
                TheExec.flow.TestLimit resultVal:=glb_TMPS_count_int, unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature" & "_" & "CoolingLoopCount", ForceResults:=tlForceNone
                TheExec.flow.TestLimit resultVal:=(Timer - Time0), unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature" & "_" & "CoolingTime", ForceResults:=tlForceNone, customUnit:="sec"
            End If
        End If
    End If
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Temperature"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Calc_Metrology_Temperature_Max_Min_Avg_Tahiti(argc As Integer, argv() As String) As Long
    'Merge Calc_MetrologyTMPS_Max_Min and Calc_MetrologyBTS_Temperature_Max_Min_Avg function -- 20220214
    'Add three of new arguments (Max/Min/Avg) to replace hardcode in this function for flexible by user define
    
    'CalcEqn: Alg:: Calc_Metrology_Temperature_Max_Min_Avg(Max=true, Min=true, Avg=true,TMP1+TMP2+TMP3)
    
    Dim Temperature As New SiteDouble
    'Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(3), "+")   'Get from argument 4th
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    
    Dim DSP_MTRSNS_Temperature_Average As New DSPWave
    Dim DSP_MTRSNS_Temperature_Average_IEDA As New DSPWave  'From T-Col
    Dim DSP_MTRSNS_Temperature_Maximum As New DSPWave
    Dim DSP_MTRSNS_Temperature_Minimum As New DSPWave
    
    'Add For T-Tah request MAX-AVG-Delta and MIN-AVG-Delta Calculations -- 20230908
    Dim DSP_MTRSNS_Temperature_Max_Avg_Delta As New DSPWave
    Dim DSP_MTRSNS_Temperature_Min_Avg_Delta As New DSPWave
    
    Dim l_DSP_MTRBTS_Temperature_Average As New DSPWave
    
    Dim Calc_MAX_FLAG As Boolean: Calc_MAX_FLAG = False 'Add Calculate Flag
    Dim Calc_MIN_FLAG As Boolean: Calc_MIN_FLAG = False 'Add Calculate Flag
    Dim Calc_AVG_FLAG As Boolean: Calc_AVG_FLAG = False 'Add Calculate Flag
    
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    Dim Sensor_Num() As String
    
    'Add Calculate Flag : Max=true
    If UCase(argv(0)) Like "*MAX*" Then
        SplitStrAry = Split(argv(0), "=")
        If UCase(SplitStrAry(1)) = "TRUE" Then
            Calc_MAX_FLAG = True
        End If
    End If
    
    'Add Calculate Flag : Min=true
    If UCase(argv(1)) Like "*MIN*" Then
        SplitStrAry = Split(argv(1), "=")
        If UCase(SplitStrAry(1)) = "TRUE" Then
            Calc_MIN_FLAG = True
        End If
    End If
    
    'Add Calculate Flag : Avg=true
    If UCase(argv(2)) Like "*AVG*" Then
        SplitStrAry = Split(argv(2), "=")
        If UCase(SplitStrAry(1)) = "TRUE" Then
            Calc_AVG_FLAG = True
        End If
    End If
    
    DSP_MTRSNS_Temperature_Average.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Average_IEDA.CreateConstant 0, 1 'From T-Col
    DSP_MTRSNS_Temperature_Maximum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Minimum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    
    'Add For T-Tah request MAX-AVG-Delta and MIN-AVG-Delta Calculations -- 20230908
    DSP_MTRSNS_Temperature_Max_Avg_Delta.CreateConstant 0, 1, DspDouble
    DSP_MTRSNS_Temperature_Min_Avg_Delta.CreateConstant 0, 1, DspDouble
    
    l_DSP_MTRBTS_Temperature_Average.CreateConstant 0, 1, DspDouble
    
    If Calc_MAX_FLAG = True Or Calc_MIN_FLAG = True Or Calc_AVG_FLAG = True Then
    
        For i = 0 To UBound(Temperature_Sensor)
            Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
            For Each site In TheExec.sites
                Temperature_Array(0) = Temperature / 64
                DSP_Temperature(i).data = Temperature_Array
                ' Update calc flag judge -- 20220214
                If Calc_AVG_FLAG = True Then DSP_MTRSNS_Temperature_Average.Element(i) = DSP_MTRSNS_Temperature_Average.Element(i) + DSP_Temperature(i).Element(0)
                    
                If Calc_MAX_FLAG = True Then DSP_MTRSNS_Temperature_Maximum.Element(i) = DSP_MTRSNS_Temperature_Maximum.Element(i) + DSP_Temperature(i).Element(0)
                If Calc_MIN_FLAG = True Then DSP_MTRSNS_Temperature_Minimum.Element(i) = DSP_MTRSNS_Temperature_Minimum.Element(i) + DSP_Temperature(i).Element(0)
            Next site
        Next i
        
        Set gS_TMPS1_Trim = New SiteVariant 'From T-Col
        
        For Each site In TheExec.sites
            ' Update calc flag judge -- 20220214
            If Calc_AVG_FLAG = True Then
                If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then '221027 for Coll change to BTS
                    'From T-Col '221027 for Coll change to BTS
                    DSP_MTRSNS_Temperature_Average_IEDA(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Average(site).CalcMean, 4)
                    DSP_MTRSNS_Temperature_Average(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Average_IEDA(site).Element(0), 0)
                    gS_TMPS1_Trim = CStr(FormatNumber(DSP_MTRSNS_Temperature_Average_IEDA(site).Element(0), 2))
                    TheExec.Datalog.WriteComment "IEDA_Temp_AVG[site" & site & "]=" & DSP_MTRSNS_Temperature_Average_IEDA(site).Element(0)
                Else
                    DSP_MTRSNS_Temperature_Average(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Average(site).CalcMean, 0)
                End If
                l_DSP_MTRBTS_Temperature_Average.Element(0) = DSP_MTRSNS_Temperature_Average(site).Element(0)
            End If
            If Calc_MAX_FLAG = True Then DSP_MTRSNS_Temperature_Maximum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Maximum(site).CalcMaximumValue, 4)
            If Calc_MIN_FLAG = True Then DSP_MTRSNS_Temperature_Minimum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Minimum(site).CalcMinimumValue, 4)
            
            
            'Add For T-Tah request MAX-AVG-Delta and MIN-AVG-Delta Calculations -- 20230908
            If Calc_MAX_FLAG = True And Calc_AVG_FLAG = True Then
                DSP_MTRSNS_Temperature_Max_Avg_Delta.Element(0) = DSP_MTRSNS_Temperature_Maximum(site).Element(0) - DSP_MTRSNS_Temperature_Average_IEDA(site).Element(0)
            End If
            If Calc_MAX_FLAG = True And Calc_AVG_FLAG = True Then
                DSP_MTRSNS_Temperature_Min_Avg_Delta.Element(0) = DSP_MTRSNS_Temperature_Minimum(site).Element(0) - DSP_MTRSNS_Temperature_Average_IEDA(site).Element(0)
            End If
        
        Next site
        Call AddStoredCaptureData(argv(argc - 1), l_DSP_MTRBTS_Temperature_Average)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , , , , , , tlForceNone)
        
        If UCase(TheExec.flow.CurrentFlowSheetName) Like "FLOW_TMPS*" Or UCase(TheExec.flow.CurrentFlowSheetName) = "FLOW_TMPS_NO_RELAY" Or UCase(TheExec.flow.CurrentFlowSheetName) = "FLOW_TMPS_TSNS" Then
            Dim TestNameInput_Ary() As String
            TestNameInput_Ary = Split(TestNameInput, "_")
            TestNameInput_Ary(8) = glb_MTRRecord & "-" & glb_MTRBTSCnt
            TestNameInput = Join(TestNameInput_Ary, "_")
        End If
        ' Update calc flag judge -- 20220214
''        If Calc_AVG_FLAG = True Then TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Average_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
''        If Calc_MAX_FLAG = True Then TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Maximum_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
''        If Calc_MIN_FLAG = True Then TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Minimum_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
      
        'From T-Col
         
      If TheExec.enableWord("HardIP_LVCC") = True Or TheExec.enableWord("HardIP_PPMN_Char") = True Then
        
            If Calc_AVG_FLAG = True Then
              If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average_IEDA.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average-D_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C"
              End If
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C"
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            End If
            If Calc_MAX_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Maximum_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C"
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            End If
            If Calc_MIN_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Minimum_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C"
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            End If
        
      Else
        If UCase(currentJobName) = "CP1" Or UCase(currentJobName) = "WLFT1" Or UCase(currentJobName) = "FT1" Then
            ' Update calc flag judge -- 20220214
            If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
                If Calc_AVG_FLAG = True Then
                  If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
                    TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average_IEDA.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average-D_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=15, hiVal:=35
                  End If
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=15, hiVal:=35
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
                If Calc_MAX_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Maximum_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=15, hiVal:=35
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
                If Calc_MIN_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Minimum_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=15, hiVal:=35
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
            Else
                If Calc_AVG_FLAG = True Then
                  If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
                    TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average_IEDA.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average-D_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=15, hiVal:=35
                  End If
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=15, hiVal:=35
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
                If Calc_MAX_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Maximum_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=15, hiVal:=35
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
                If Calc_MIN_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Minimum_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=15, hiVal:=35
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
            
            End If
            

        ElseIf UCase(currentJobName) = "CP2" Or UCase(currentJobName) = "FT2" Then
            ' Update calc flag judge -- 20220214
            If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
                If Calc_AVG_FLAG = True Then
                  If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
                    TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average_IEDA.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average-D_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=70, hiVal:=100
                  End If
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=70, hiVal:=100
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
                If Calc_MAX_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Maximum_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=70, hiVal:=100
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
                If Calc_MIN_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Minimum_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=70, hiVal:=100
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
            Else
                If Calc_AVG_FLAG = True Then
                  If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
                    TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average_IEDA.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average-D_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", lowVal:=70, hiVal:=100
                  End If
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=70, hiVal:=100
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
                If Calc_MAX_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Maximum_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=70, hiVal:=100
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
                If Calc_MIN_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Minimum_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=70, hiVal:=100
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                End If
            
            End If
            

        Else
            ' Update calc flag judge -- 20220214
            If Calc_AVG_FLAG = True Then
              If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average_IEDA.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average-D_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=15, hiVal:=100
              End If
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Average_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=15, hiVal:=100
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            End If
            If Calc_MAX_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Maximum_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=15, hiVal:=100
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            End If
            If Calc_MIN_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Minimum_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", lowVal:=15, hiVal:=100
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            End If
        End If
      End If
    End If
    
 If UCase(TheExec.DataManager.instancename) Like "MTRBTS_M1TMP10T4P1C_PP*" Then
    'Add For T-Tah request MAX-AVG-Delta and MIN-AVG-Delta Calculations -- 20230908
    If UCase(currentJobName) = "CP1" Then
        If Calc_MAX_FLAG = True And Calc_MAX_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Max_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Max-Avg_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", scaletype:=scaleNoScaling, lowVal:=0, hiVal:=1.5
        End If
        If Calc_MAX_FLAG = True And Calc_MIN_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Min_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Min-Avg_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", scaletype:=scaleNoScaling, lowVal:=-1.5, hiVal:=0
        End If
    ElseIf UCase(currentJobName) = "CP2" Then
        If Calc_MAX_FLAG = True And Calc_MAX_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Max_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Max-Avg_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", scaletype:=scaleNoScaling, lowVal:=0, hiVal:=5
        End If
        If Calc_MAX_FLAG = True And Calc_MIN_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Min_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Min-Avg_x_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C", scaletype:=scaleNoScaling, lowVal:=-5, hiVal:=0
        End If
    ElseIf UCase(currentJobName) = "FT1" Or UCase(currentJobName) = "WLFT1" Then
        If Calc_MAX_FLAG = True And Calc_MAX_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Max_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Max-Avg_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", scaletype:=scaleNoScaling, lowVal:=0, hiVal:=3
        End If
        If Calc_MAX_FLAG = True And Calc_MIN_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Min_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Min-Avg_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", scaletype:=scaleNoScaling, lowVal:=-3, hiVal:=0
        End If
    ElseIf UCase(currentJobName) = "FT2" Then
        If Calc_MAX_FLAG = True And Calc_MAX_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Max_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Max-Avg_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", scaletype:=scaleNoScaling, lowVal:=0, hiVal:=10
        End If
        If Calc_MAX_FLAG = True And Calc_MIN_FLAG = True Then
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Min_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature-Min-Avg_x_", , , vbTextCompare), ForceResults:=tlForceNone, customUnit:="C", scaletype:=scaleNoScaling, lowVal:=-10, hiVal:=0
        End If
    End If
 End If
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_Metrology_Temperature_Max_Min_Avg"
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function Calc_Metrology_Temperature_Max_Min_Avg(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    'Merge Calc_MetrologyTMPS_Max_Min and Calc_MetrologyBTS_Temperature_Max_Min_Avg function -- 20220214
    'Add three of new arguments (Max/Min/Avg) to replace hardcode in this function for flexible by user define
    
    'CalcEqn: Alg:: Calc_Metrology_Temperature_Max_Min_Avg(Max=true, Min=true, Avg=true,TMP1+TMP2+TMP3)
    
    Dim Temperature As New SiteDouble
    'Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(3), "+")   'Get from argument 4th
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature As New DSPWave
    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    
    Dim DSP_MTRSNS_Temperature_Average As New DSPWave
    Dim DSP_MTRSNS_Temperature_Maximum As New DSPWave
    Dim DSP_MTRSNS_Temperature_Minimum As New DSPWave
    Dim DSP_MTRSNS_Temperature_Maximum_TMPS_MON As New DSPWave
    Dim DSP_MTRSNS_Temperature_Max_Avg_Delta As New DSPWave
    Dim DSP_MTRSNS_Temperature_Min_Avg_Delta As New DSPWave
    
    Dim l_DSP_MTRBTS_Temperature_Average As New DSPWave 'Add to store average temparature
    
    Dim Calc_MAX_FLAG As Boolean: Calc_MAX_FLAG = False 'Add Calculate Flag
    Dim Calc_MIN_FLAG As Boolean: Calc_MIN_FLAG = False 'Add Calculate Flag
    Dim Calc_AVG_FLAG As Boolean: Calc_AVG_FLAG = False 'Add Calculate Flag
    Dim SplitStrAry() As String
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    Dim Sensor_Num() As String
    
    Dim b_TMPS_MON_Flow As Boolean
    
    If InStr(Instance_Data.CUS_Str_DigSrcData, "TMPS_MON") > 0 Then
        b_TMPS_MON_Flow = True
    Else
        b_TMPS_MON_Flow = False
    End If
    
    'Add Calculate Flag : Max=true
    If UCase(argv(0)) Like "*MAX*" Then
        SplitStrAry = Split(argv(0), "=")
        If UCase(SplitStrAry(1)) = "TRUE" Then
            Calc_MAX_FLAG = True
        End If
    End If
    
    'Add Calculate Flag : Min=true
    If UCase(argv(1)) Like "*MIN*" Then
        SplitStrAry = Split(argv(1), "=")
        If UCase(SplitStrAry(1)) = "TRUE" Then
            Calc_MIN_FLAG = True
        End If
    End If
    
    'Add Calculate Flag : Avg=true
    If UCase(argv(2)) Like "*AVG*" Then
        SplitStrAry = Split(argv(2), "=")
        If UCase(SplitStrAry(1)) = "TRUE" Then
            Calc_AVG_FLAG = True
        End If
    End If
    
    DSP_MTRSNS_Temperature.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Average.CreateConstant 0, 1
    DSP_MTRSNS_Temperature_Maximum.CreateConstant 0, 1
    DSP_MTRSNS_Temperature_Minimum.CreateConstant 0, 1
    DSP_MTRSNS_Temperature_Max_Avg_Delta.CreateConstant 0, 1
    DSP_MTRSNS_Temperature_Min_Avg_Delta.CreateConstant 0, 1
    DSP_MTRSNS_Temperature_Maximum_TMPS_MON.CreateConstant 0, 1
    l_DSP_MTRBTS_Temperature_Average.CreateConstant 0, 1, DspDouble
    
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).data = Temperature_Array
            DSP_MTRSNS_Temperature.Element(i) = DSP_MTRSNS_Temperature.Element(i) + DSP_Temperature(i).Element(0)
        Next site
    Next i
    
    If b_TMPS_MON_Flow Then
        glb_TMPS_fail_flag = False
        
        For Each site In TheExec.sites
            DSP_MTRSNS_Temperature_Maximum_TMPS_MON(site).Element(0) = DSP_MTRSNS_Temperature(site).CalcMaximumValue
            glb_TMPS_start_flag = TheExec.sites.item(site).FlagState("F_TMPS_Start")
            glb_TMPS_end_flag = TheExec.sites.item(site).FlagState("F_TMPS_End")
            
            If Instance_Data.hiLimit(TheExec.flow.TestLimitIndex) = "" Then
            ElseIf DSP_MTRSNS_Temperature_Maximum_TMPS_MON(site).Element(0) > CDbl(Instance_Data.hiLimit(TheExec.flow.TestLimitIndex)) Then
                glb_TMPS_fail_flag = True
            End If
        Next site
        
        Dim TMPS_timer As Variant
        If TheExec.enableWord("Enable_Vddbinning_Cooling") = False Then
            glb_TMPS_count_str = "Last"
            glb_TMPS_count_int = 0
            TMPS_timer = 0
        Else
            If glb_TMPS_start_flag = True Then
                glb_TMPS_count_str = "First"
                glb_TMPS_count_int = 0
                glb_TMPS_timer = Timer
            ElseIf glb_TMPS_fail_flag = False Or glb_TMPS_end_flag = True Then
                glb_TMPS_count_str = "Last"
                glb_TMPS_count_int = glb_TMPS_count_int + 1
            Else
                glb_TMPS_count_int = glb_TMPS_count_int + 1
                glb_TMPS_count_str = CStr(glb_TMPS_count_int)
            End If
        End If
    End If
    
    For i = 0 To UBound(Temperature_Sensor)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        If b_TMPS_MON_Flow Then
            TestNameInput = "HAC_MTRBTS_" & glb_TestInstance_TMPS & "_" & Replace(Split(TestNameInput, "_")(6), "-", "_") & "_" & glb_TMPS_count_str
        End If
        TheExec.flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
    If Calc_MAX_FLAG = True Or Calc_MIN_FLAG = True Or Calc_AVG_FLAG = True Then
        For Each site In TheExec.sites
            ' Update calc flag judge -- 20220214
            If Calc_AVG_FLAG = True Then
'                DSP_MTRSNS_Temperature_Average(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature(site).CalcMean, 0)
                DSP_MTRSNS_Temperature_Average(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature(site).CalcMean, 4)
                l_DSP_MTRBTS_Temperature_Average.Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Average(site).Element(0), 0)
            End If
            If Calc_MAX_FLAG = True Then DSP_MTRSNS_Temperature_Maximum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature(site).CalcMaximumValue, 4)
            If Calc_MIN_FLAG = True Then DSP_MTRSNS_Temperature_Minimum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature(site).CalcMinimumValue, 4)
            
            If Calc_MAX_FLAG = True And Calc_MIN_FLAG = True And Calc_AVG_FLAG = True Then
                DSP_MTRSNS_Temperature_Max_Avg_Delta(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Maximum(site).Element(0) - DSP_MTRSNS_Temperature_Average(site).Element(0), 4)
                DSP_MTRSNS_Temperature_Min_Avg_Delta(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Minimum(site).Element(0) - DSP_MTRSNS_Temperature_Average(site).Element(0), 4)
            End If
        Next site
        
        If b_TMPS_MON_Flow Then
        
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance_TMPS & "_X_Temperature_Average" & "_" & glb_TMPS_count_str, ForceResults:=tlForceNone, customUnit:="C"
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance_TMPS & "_X_Temperature_Maximum" & "_" & glb_TMPS_count_str, ForceResults:=tlForceNone, customUnit:="C"
            TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance_TMPS & "_X_Temperature_Minimum" & "_" & glb_TMPS_count_str, ForceResults:=tlForceNone, customUnit:="C"
            
        Else
            '20230203 add for MTRBTS store avg temparature
            Call StoreDataAllType(argv(argc - 1), l_DSP_MTRBTS_Temperature_Average)
            
            TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , , , , , , tlForceNone)
            
            ' Update calc flag judge -- 20220214
            If Calc_AVG_FLAG = True Then TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Average_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
            If Calc_MAX_FLAG = True Then TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Maximum_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
            If Calc_MIN_FLAG = True Then TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Minimum_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
        
            If Calc_MAX_FLAG = True And Calc_MIN_FLAG = True And Calc_AVG_FLAG = True Then
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Max_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Max-Avg-Delta_", , , vbTextCompare), ForceResults:=tlForceFlow
                TheExec.flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Min_Avg_Delta.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Min-Avg-Delta_", , , vbTextCompare), ForceResults:=tlForceFlow
            End If
        End If
    End If
    
    If b_TMPS_MON_Flow = True And TheExec.enableWord("Enable_Vddbinning_Cooling") = True And (glb_TMPS_fail_flag = False Or glb_TMPS_end_flag = True) Then
        If glb_TMPS_start_flag = True Then
            TMPS_timer = 0
        Else
            TMPS_timer = Timer - glb_TMPS_timer
        End If
        
        TheExec.flow.TestLimit resultVal:=glb_TMPS_count_int, unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature" & "_" & "CoolingLoopCount", ForceResults:=tlForceNone
        TheExec.flow.TestLimit resultVal:=TMPS_timer, unit:=unitCustom, Tname:="HAC_MTRBTS_" & glb_TestInstance & "_X_Temperature" & "_" & "CoolingTime", ForceResults:=tlForceNone, customUnit:="sec"
    End If
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_Metrology_Temperature_Max_Min_Avg") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MetrologyBTS_Temperature_DOE(argc As Integer, argv() As String) As Long
    'From T-Col
    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    Dim temp_instance_name As String
    Dim Src_count As Long
    Dim Temp_ForLoopIntegerName() As String
    temp_instance_name = TheExec.DataManager.instancename
    'If InStr(Temp_instance_Name, "_NV") = 0 Then
        Temp_ForLoopIntegerName = Split(Instance_Data.DigSrc_FlowForLoopIntegerName, ";")
        Src_count = TheExec.flow.var(Temp_ForLoopIntegerName(0)).value
    'End If

    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoredData(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature = Format(Temperature.divide(64), "0.0000")
            Temperature_Array(0) = Temperature
            DSP_Temperature(i).data = Temperature_Array
        Next site
        'If InStr(Temp_instance_Name, "_NV") = 0 Then
            Call AddStoredData(Temperature_Sensor(i) + "_" + CStr(Src_count), Temperature)
        'End If
        Set Temperature = Nothing
    Next i
    
    For i = 0 To UBound(Temperature_Sensor)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Temperature"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Calc_MetrologyBTS_Average_And_Stdev_Temperature(argc As Integer, argv() As String) As Long
    'From T-Col
    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim Average_Temperature() As New DSPWave: ReDim Average_Temperature(UBound(Temperature_Sensor))
    Dim Stdev_Temperature() As New DSPWave: ReDim Stdev_Temperature(UBound(Temperature_Sensor))
    Dim Temp_ForLoopIntegerName() As String
    Dim Src_count As Long
    Dim Temp_Temperature As New SiteDouble
    Dim Temp_Average As New SiteDouble
    Dim Temp_Average_Tempertaure As New SiteDouble
    
    Dim i, j As Integer
    Dim TestNameInput As String
    
    'Temp_ForLoopIntegerName = Split(Instance_Data.DigSrc_FlowForLoopIntegerName, ";")
    'Src_count = theexec.Flow.var(Temp_ForLoopIntegerName(0)).Value
    Src_count = TheExec.flow.var("SrcCodeIndx").value

    If Src_count = 15 Then
        For i = 0 To UBound(Temperature_Sensor)
            For j = 0 To 15
                Temp_Temperature = GetStoredData(Temperature_Sensor(i) + "_" + CStr(j))
                Temperature = Temperature.Add(Temp_Temperature.divide(64))
            Next j
            For Each site In TheExec.sites
                Temp_Average_Tempertaure = Temperature.divide(16)
                Temperature_Array(0) = Temp_Average_Tempertaure
                Average_Temperature(i).data = Temperature_Array
            Next site
            Temperature = 0
            For j = 0 To 15
                Temp_Temperature = GetStoredData(Temperature_Sensor(i) + "_" + CStr(j))
                Temp_Average = Temp_Temperature.divide(64).Subtract(Temp_Average_Tempertaure).power(2)
                Temperature = Temperature.Add(Temp_Average)
            Next j
            For Each site In TheExec.sites
                Temperature_Array(0) = Temperature.divide(16).power(0.5)
                Stdev_Temperature(i).data = Temperature_Array
            Next site
        Next i
    
    For i = 0 To UBound(Temperature_Sensor)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Average_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Stdev_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i

    End If
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Temperature"
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Calc_MetrologyBTS_Average_And_Stdev_Temperature_DOE(argc As Integer, argv() As String) As Long
    'From T-Col
    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim Average_Temperature() As New DSPWave: ReDim Average_Temperature(UBound(Temperature_Sensor))
    Dim Stdev_Temperature() As New DSPWave: ReDim Stdev_Temperature(UBound(Temperature_Sensor))
    Dim Temp_ForLoopIntegerName() As String
    Dim Src_count As Long
    Dim Temp_Temperature As New SiteDouble
    Dim Temp_Average As New SiteDouble
    Dim Temp_Average_Tempertaure As New SiteDouble
    Temp_ForLoopIntegerName = Split(Instance_Data.DigSrc_FlowForLoopIntegerName, ";")
    Src_count = TheExec.flow.var(Temp_ForLoopIntegerName(0)).value

    If Src_count = 15 Then
        For i = 0 To UBound(Temperature_Sensor)
            For j = 0 To 15
                Temp_Temperature = GetStoredData(Temperature_Sensor(i) + "_" + CStr(j))
                Temperature = Temperature.Add(Temp_Temperature)
            Next j
            For Each site In TheExec.sites
                Temp_Average_Tempertaure = Format(Temperature.divide(16), "0.0000")
                Temperature_Array(0) = Temp_Average_Tempertaure
                Average_Temperature(i).data = Temperature_Array
            Next site
            Temperature = 0
            For j = 0 To 15
                Temp_Temperature = GetStoredData(Temperature_Sensor(i) + "_" + CStr(j))
                Temp_Average = Temp_Temperature.Subtract(Temp_Average_Tempertaure).power(2)
                Temperature = Temperature.Add(Temp_Average)
            Next j
            For Each site In TheExec.sites
                Temperature_Array(0) = Temperature.divide(16).power(0.5)
                Stdev_Temperature(i).data = Temperature_Array
            Next site
            Temperature = 0
        Next i
    
    For i = 0 To UBound(Temperature_Sensor)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Average_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Stdev_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i

    End If
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Temperature"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Calc_Metrology_nTS_T1(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'Calc_Metrology_nTS_T1(VREF_nTS_G,VREFN_nTS_G,nts_vref_dig_G_cp1)
    Dim Dict_V1 As String
    Dim Dict_V2 As String
    Dim Dict_Vdiff As String
    Dim Dict_nts_vref_dig As String
    
    Dim SD_V1 As New PinListData
    Dim SD_V2 As New PinListData
    Dim SD_Vdiff As New PinListData
    Dim SD_Vstep1 As New PinListData
    Dim SD_nts_vref_dig As New PinListData
    Dim TestNameInput As String
    Dim site As Variant
    Dim DSPW_SD_nts_vref_dig As New DSPWave
    Dim DSPW_SD_nts_vref_dig_Bin As New DSPWave
    Dim temp(0) As Double

    
    DSPW_SD_nts_vref_dig.CreateConstant 0, 12, DspLong
    DSPW_SD_nts_vref_dig_Bin.CreateConstant 0, 12, DspLong
    
    Dict_V1 = argv(0)
    Dict_V2 = argv(1)
    Dict_nts_vref_dig = argv(2)
    

'    If TheExec.TesterMode = testModeOffline Then            'for offline run
'    SD_V1.AddPin ("MTR_ATB")
'    SD_V2.AddPin ("MTR_ATB")
'        For Each Site In TheExec.sites
'            SD_V1.Pins(0) = 1.5 + Site * 0.01
'            SD_V2.Pins(0) = 1.2 + Site * 0.03
'        Next Site
'    Else

    ''' Update for new NTS pattern @William 230407
    If InStr(Dict_V1, "_para") <> 0 And InStr(Dict_V2, "_para") Then
        SD_V1.AddPin ("Capture_Pin")
        SD_V2.AddPin ("Capture_Pin")
        
        SD_V1.pins(0) = GetStoreDataAllType(Dict_V1)
        SD_V2.pins(0) = GetStoreDataAllType(Dict_V2)
    Else
        SD_V1 = GetStoreDataAllType(Dict_V1)
        SD_V2 = GetStoreDataAllType(Dict_V2) 'Remove for V09A
    End If
'
'    End If
    
    'SD_Vdiff = SD_V1.Math.Subtract(SD_V2)
    SD_Vdiff = SD_V1 ' For V09A
    
    'SD_Vstep1 = SD_Vdiff.Math.divide(0.4)
    SD_Vstep1 = SD_Vdiff.Math.divide(0.5) ' For V09A
    SD_nts_vref_dig = SD_Vstep1.Math.Multiply(2048)
    For Each site In TheExec.sites
        SD_nts_vref_dig = Round(SD_nts_vref_dig)
    Next site
    
    Call StoreDataAllType("PLD_" & Dict_nts_vref_dig, SD_nts_vref_dig)
    
    For Each site In TheExec.sites
        temp(0) = SD_nts_vref_dig.pins(0).value
        DSPW_SD_nts_vref_dig.data = temp
    Next site
    Call rundsp.DSPWaveDecToBinary(DSPW_SD_nts_vref_dig, 12, DSPW_SD_nts_vref_dig_Bin)
    Call StoreDataAllType(Dict_nts_vref_dig, DSPW_SD_nts_vref_dig_Bin)

    If gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment "====Calc Function: Calc_Metrology_nTS_T1 debug_log_Start===="
        For Each site In TheExec.sites
            'theexec.Datalog.WriteComment "site_" & site & "_Vref_Differential_nTS : " & SD_Vdiff
            TheExec.Datalog.WriteComment "site_" & site & "_Vref_nTS : " & SD_Vdiff 'For V09A
            TheExec.Datalog.WriteComment "site_" & site & "_nts_vref_dig          : " & SD_nts_vref_dig
        Next site
        TheExec.Datalog.WriteComment "====Calc Function: Calc_Metrology_nTS_T1 debug_log_End===="
    End If
    TestNameInput = Report_TName_From_Instance(CalcV, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=SD_nts_vref_dig.pins(0), Tname:=TestNameInput, ForceResults:=tlForceFlow   'tlForceFlow
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_Metrology_nTS_T1") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
    
End Function



Public Function Calc_MetrologyHSCnTS_STD_Coefficient_N3(argc As Long, argv() As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'Calc_MetrologyHSCnTS_STD_Coefficient_N3(
'hsc_nts_tp00g_d_0_cp1,
'hsc_nts_tp00g_vref_dig_0_cp1,
'nts_vref_dig_G_ft2,25,
'hsc_nts_tp00g_c0,
'hsc_nts_tp00g_c1,
'hsc_nts_tp00g_c2)

    Dim counter As Integer
    Dim str1 As String
    Dim str2 As String
        
    Dim site As Variant
    Dim i As Integer
    Dim Dict_ReadFromEfuse As String
    Dict_ReadFromEfuse = argv(0)
    
    Dim ReadEfuse_Bin As New DSPWave
    Dim ReadEfuse_Dec As New DSPWave
    'ReadEfuse_Bin.CreateConstant 0, 21, DspLong    '20240131 Hidra by YM  Dout 17Bit
    ReadEfuse_Dec.CreateConstant 0, 1, DspLong
    
    Dim DSPWF_LSB As New DSPWave
    Dim DSPWF_MSB As New DSPWave
    Dim DSPWF_Mid As New DSPWave
    '20240131 Hidra by YM  Dout 17Bit
    'DSPWF_LSB.CreateConstant 0, 16, DspLong
    'DSPWF_MSB.CreateConstant 0, 1, DspLong
    'DSPWF_Mid.CreateConstant 0, 4, DspLong
    
    Dim x As New SiteLong
    
    
    If TheExec.TesterMode = testModeOffline Then            'for offline run

    'readEfuse=LSB------------------MSB
        str1 = "01111001111110010"
        str2 = "11111000000110111"
        ReadEfuse_Bin.CreateConstant 0, Len(str1), DspLong  '20240131 Hidra by YM  Dout 17Bit
        For Each site In TheExec.sites
            If site = 0 Then
                For i = 0 To Len(str1) - 1
                    ReadEfuse_Bin.Element(i) = mid(str1, i + 1, 1)
                Next i
            Else
                For i = 0 To Len(str2) - 1
                    ReadEfuse_Bin.Element(i) = mid(str2, i + 1, 1)
                Next i
            End If
        Next site
    Else
        ReadEfuse_Bin = GetStoreDataAllType(argv(0))
    End If
    '20240131 Hidra by YM  Dout 17Bit
'    For Each site In TheExec.sites
'        x = ReadEfuse_Bin.SampleSize - 1
'        For i = 0 To x
'            If i < 16 Then
'                DSPWF_LSB.Element(i) = ReadEfuse_Bin.Element(i)
'            ElseIf i = x Then
'                DSPWF_MSB.Element(i - x) = ReadEfuse_Bin.Element(i)
'            ElseIf i > 15 And i < 21 Then
'                DSPWF_Mid.Element(i - 16) = ReadEfuse_Bin.Element(i)
'            End If
'        Next i
'    Next site

    DSPWF_LSB.CreateConstant 0, 16, DspLong
    DSPWF_MSB.CreateConstant 0, 1, DspLong
    DSPWF_Mid.CreateConstant 0, 4, DspLong
    'LSB [15:0]
    rundsp.DSP_SelectBits ReadEfuse_Bin, DSPWF_LSB, 0, 16
    'MSB [16:16]
    rundsp.DSP_SelectBits ReadEfuse_Bin, DSPWF_MSB, 16, 1
    '20240131 Hidra by YM  Dout 17Bit
    
    Call rundsp.BinToDec(DSPWF_LSB, ReadEfuse_Dec)
    For Each site In TheExec.sites
        ''' Update for Donan @William 230118
        If DSPWF_MSB.Element(0) = 0 Then
            ReadEfuse_Dec = ReadEfuse_Dec.divide(2 ^ 16)
        Else
            ReadEfuse_Dec = ReadEfuse_Dec.divide(2 ^ 16)
            ReadEfuse_Dec = ReadEfuse_Dec.Subtract(1)
        End If
    Next site
   
'    Call StoreDataAllType(Dict_ReadFromEfuse, ReadEfuse_Dec)
    
    If gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment "====Calc Function: Calc_MetrologyHSCnTS_Coefficient_N3 debug_log_Start===="
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site_" & site & "_nts_d_" & right(Split(Dict_ReadFromEfuse, "_")(2), 1) & "_cp1_Dec : " & ReadEfuse_Dec.Element(0)
        Next site
    End If
    
'=========step2========================
    Dim ReadEfuse2_cp1_Bin As New DSPWave
    Dim ReadEfuse2_ft2_Bin As New DSPWave
    Dim Gain_cp1_Dec As New DSPWave
    Dim Gain_ft2_Dec As New DSPWave
    Dim str3 As String
    Dim str4 As String
    Dim Dict_Gain_cp1 As String
    Dim Dict_Gain_ft2 As String
    
    Dict_Gain_cp1 = argv(1) & "_Gain"
    Dict_Gain_ft2 = argv(2) & "_Gain"
    
    Dim y As New SiteLong
    
    ReadEfuse2_cp1_Bin.CreateConstant 0, 12, DspLong
    ReadEfuse2_ft2_Bin.CreateConstant 0, 12, DspLong
    Gain_cp1_Dec.CreateConstant 0, 1, DspLong
    
    If TheExec.TesterMode = testModeOffline Then            'for offline run

    'readEfuse=LSB----------MSB
        str3 = "000000000001"
        str4 = "101001111010"
        
        For Each site In TheExec.sites
            For i = 0 To Len(str3) - 1
                ReadEfuse2_cp1_Bin.Element(i) = mid(str3, i + 1, 1)
            Next i
            For i = 0 To Len(str4) - 1
                ReadEfuse2_ft2_Bin.Element(i) = mid(str4, i + 1, 1)
            Next i
        Next site
    Else
        ReadEfuse2_cp1_Bin = GetStoreDataAllType(argv(1))
        ReadEfuse2_ft2_Bin = GetStoreDataAllType(argv(2))
    End If
    
    Dim DSPWF2_cp1_Data As New DSPWave
    Dim DSPWF2_ft2_Data As New DSPWave
    DSPWF2_cp1_Data.CreateConstant 0, 12, DspLong
    DSPWF2_ft2_Data.CreateConstant 0, 12, DspLong
    
    For Each site In TheExec.sites
        y = ReadEfuse2_cp1_Bin.SampleSize - 1
        For i = 0 To y
            DSPWF2_cp1_Data.Element(i) = ReadEfuse2_cp1_Bin.Element(i)
            DSPWF2_ft2_Data.Element(i) = ReadEfuse2_ft2_Bin.Element(i)
        Next i
    Next site
    
    Call rundsp.BinToDec(DSPWF2_cp1_Data, Gain_cp1_Dec)
    Call rundsp.BinToDec(DSPWF2_ft2_Data, Gain_ft2_Dec)
    For Each site In TheExec.sites
        Gain_cp1_Dec = Gain_cp1_Dec.divide(2048)
        Gain_ft2_Dec = Gain_ft2_Dec.divide(2048)
    Next site
    
    Call StoreDataAllType(Dict_Gain_cp1, Gain_cp1_Dec)
    Call StoreDataAllType(Dict_Gain_ft2, Gain_ft2_Dec)
    
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site_" & site & "_Gain_cp1_Dec : " & Gain_cp1_Dec.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_Gain_ft2_Dec : " & Gain_ft2_Dec.Element(0)
        Next site
    End If
'=========step3========================
    Dim a0_gold_0 As New SiteDouble
    Dim a1_gold_0 As New SiteDouble
    Dim a2_gold_0 As New SiteDouble
    Dim a0_gold As New DSPWave
    Dim a1_gold As New DSPWave
    Dim a2_gold As New DSPWave
    Dim Dict_nts_c0_i As String
    Dim Dict_nts_c1_i As String
    Dim Dict_nts_c2_i As String

    Dim k_i As New DSPWave
    Dim k_i_1 As New DSPWave

    Dim C0_i As New DSPWave
    Dim C1_i As New DSPWave
    Dim C2_i As New DSPWave

    Dim C0_i_1 As New DSPWave
    Dim C0_i_2 As New DSPWave

    Dim C1_i_1 As New DSPWave
    Dim C1_i_2 As New DSPWave

    Dim C_wdth As New SiteLong
    Dim C0_i_2S_C As New DSPWave
    Dim C1_i_2S_C As New DSPWave
    Dim C2_i_2S_C As New DSPWave
    Dim C0_i_2S_C_Bin As New DSPWave
    Dim C1_i_2S_C_Bin As New DSPWave
    Dim C2_i_2S_C_Bin As New DSPWave
    Dim Dict_C0_i_2S_C As String
    Dim Dict_C1_i_2S_C As String
    Dim Dict_C2_i_2S_C As String
    Dim TestNameInput As String
    Dim BKM_DECODE As String
    Dim sBKM As String      'Update for parsing MTR table -- 20230103
    
    Dict_C0_i_2S_C = argv(4)
    Dict_C1_i_2S_C = argv(5)
    Dict_C2_i_2S_C = argv(6)
    
    a0_gold.CreateConstant 0, 1, DspLong
    a1_gold.CreateConstant 0, 1, DspLong
    a2_gold.CreateConstant 0, 1, DspLong
    k_i.CreateConstant 0, 1, DspLong
    k_i_1.CreateConstant 0, 1, DspLong
    C0_i.CreateConstant 0, 1, DspLong
    C1_i.CreateConstant 0, 1, DspLong
    C2_i.CreateConstant 0, 1, DspLong
    C0_i_1.CreateConstant 0, 1, DspLong
    C0_i_2.CreateConstant 0, 1, DspLong
    C1_i_1.CreateConstant 0, 1, DspLong
    C1_i_2.CreateConstant 0, 1, DspLong
    C0_i_2S_C.CreateConstant 0, 1, DspLong
    C1_i_2S_C.CreateConstant 0, 1, DspLong
    C2_i_2S_C.CreateConstant 0, 1, DspLong
    C0_i_2S_C_Bin.CreateConstant 0, 25, DspLong
    C1_i_2S_C_Bin.CreateConstant 0, 25, DspLong
    C2_i_2S_C_Bin.CreateConstant 0, 25, DspLong

'    If TheExec.TesterMode = testModeOffline Then BKM_DECODE = 2

'    If BKM_DECODE = "1" Or BKM_DECODE = "8" Then
'        a0_gold_0 = 3.22      (0x0000671, Q16.9 format)
'        a1_gold_0 = 470.75    (0x003ad80)
'        a2_gold_0 = -175.2    (0x1fea19a)
'        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM 2.1 *V3* ) **********************"
'        TheExec.Datalog.WriteComment "  a0_gold_0 = 3.22 (0x0000671)"
'        TheExec.Datalog.WriteComment "  a1_gold_0 = 470.75 (0x003ad80)"
'        TheExec.Datalog.WriteComment "  a2_gold_0 = -175.2 (0x1fea19a)"
'        TheExec.Datalog.WriteComment "*******************************************************************************************"

'    ElseIf BKM_DECODE = "2" Or BKM_DECODE = "9" Then
'        a0_gold_0 = 3.22      (0x0000671, Q16.9 format)
'        a1_gold_0 = 470.75    (0x003ad80)
'        a2_gold_0 = -175.2    (0x1fea19a)
'        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM 2.2 *V3* ) **********************"
'        TheExec.Datalog.WriteComment "  a0_gold_0 = 3.22 (0x0000671)"
'        TheExec.Datalog.WriteComment "  a1_gold_0 = 470.75 (0x003ad80)"
'        TheExec.Datalog.WriteComment "  a2_gold_0 = -175.2 (0x1fea19a)"
'        TheExec.Datalog.WriteComment "*******************************************************************************************"

'    ElseIf BKM_DECODE = "3" Or BKM_DECODE = "10" Then
'        a0_gold_0 = 3.22      (0x0000671, Q16.9 format)
'        a1_gold_0 = 470.75    (0x003ad80)
'        a2_gold_0 = -175.2    (0x1fea19a)
'        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM 2.3 *V3* ) **********************"
'        TheExec.Datalog.WriteComment "  a0_gold_0 = 3.22 (0x0000671)"
'        TheExec.Datalog.WriteComment "  a1_gold_0 = 470.75 (0x003ad80)"
'        TheExec.Datalog.WriteComment "  a2_gold_0 = -175.2 (0x1fea19a)"
'        TheExec.Datalog.WriteComment "*******************************************************************************************"
'    Else
    a0_gold_0 = 1.96
    a1_gold_0 = 467
    a2_gold_0 = -141.26
    TheExec.Datalog.WriteComment "********************** Coefficients for Calibration **********************"
    TheExec.Datalog.WriteComment "  a0_gold_0 = " & a0_gold_0
    TheExec.Datalog.WriteComment "  a1_gold_0 = " & a1_gold_0
    TheExec.Datalog.WriteComment "  a2_gold_0 = " & a2_gold_0
    TheExec.Datalog.WriteComment "**************************************************************************"
         
    Dim tmp0(0) As Double
    Dim tmp1(0) As Double
    Dim tmp2(0) As Double

    For Each site In TheExec.sites
        
        C_wdth(site) = argv(3)
        tmp0(0) = a0_gold_0
        tmp1(0) = a1_gold_0
        tmp2(0) = a2_gold_0

        a0_gold.data = tmp0
        a1_gold.data = tmp1
        a2_gold.data = tmp2

        'k_i = Gain_ft2_Dec.Multiply(ReadEfuse_Dec)
        k_i = Gain_cp1_Dec.Multiply(ReadEfuse_Dec) ' For V09A
        k_i_1 = Gain_ft2_Dec.Subtract(Gain_cp1_Dec)
        k_i_1 = k_i_1.divide(0.6)
        k_i = k_i.Add(k_i_1)

        a0_gold = a0_gold.Add(25)
        C0_i_1 = k_i.Multiply(a1_gold)
        C0_i_2 = k_i.Multiply(k_i)
        C0_i_2 = C0_i_2.Multiply(a2_gold)

        C0_i = C0_i_1.Add(C0_i_2)
        C0_i = C0_i.Add(a0_gold)

        C1_i_1 = Gain_ft2_Dec.Multiply(a1_gold)
        C1_i_1 = C1_i_1.Multiply(-1)
        C1_i_2 = Gain_ft2_Dec.Multiply(k_i)
        C1_i_2 = C1_i_2.Multiply(a2_gold)
        C1_i_2 = C1_i_2.Multiply(-2)

        C1_i = C1_i_1.Add(C1_i_2)

        C2_i = Gain_ft2_Dec.Multiply(Gain_ft2_Dec)
        C2_i = C2_i.Multiply(a2_gold)
        
        C0_i = C0_i.Multiply(512)
        C1_i = C1_i.Multiply(512)
        C2_i = C2_i.Multiply(512)
        
    Next site
            
    Call rundsp.DSP_Convert_2S_Complement(C0_i, C_wdth, C0_i_2S_C)
    Call rundsp.DSP_Convert_2S_Complement(C1_i, C_wdth, C1_i_2S_C)
    Call rundsp.DSP_Convert_2S_Complement(C2_i, C_wdth, C2_i_2S_C)

    Call rundsp.DSPWaveDecToBinary(C0_i_2S_C, C_wdth, C0_i_2S_C_Bin)
    Call rundsp.DSPWaveDecToBinary(C1_i_2S_C, C_wdth, C1_i_2S_C_Bin)
    Call rundsp.DSPWaveDecToBinary(C2_i_2S_C, C_wdth, C2_i_2S_C_Bin)

    Call StoreDataAllType(Dict_C0_i_2S_C, C0_i_2S_C_Bin)
    Call StoreDataAllType(Dict_C1_i_2S_C, C1_i_2S_C_Bin)
    Call StoreDataAllType(Dict_C2_i_2S_C, C2_i_2S_C_Bin)
    
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites

            TheExec.Datalog.WriteComment "site_" & site & "_k_i : " & k_i(site).Element(0) & " " & "site_" & site & "_nts_c0_" & right(Split(Dict_C0_i_2S_C, "_")(2), 1) & "_Dec : " & C0_i_2S_C.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_k_i : " & k_i(site).Element(0) & " " & "site_" & site & "_nts_c1_" & right(Split(Dict_C1_i_2S_C, "_")(2), 1) & "_Dec : " & C1_i_2S_C.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_k_i : " & k_i(site).Element(0) & " " & "site_" & site & "_nts_c2_" & right(Split(Dict_C2_i_2S_C, "_")(2), 1) & "_Dec : " & C2_i_2S_C.Element(0)
        Next site
        TheExec.Datalog.WriteComment "====Calc Function: Calc_MetrologyHSCnTS_Coefficient_N3 debug_log_End===="
    End If

    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=C0_i_2S_C.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=C1_i_2S_C.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=C2_i_2S_C.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyHSCnTS_STD_Coefficient_N3") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

End Function



Public Function Calc_MetrologyHSCnTS_ADV_Coefficient(argc As Long, argv() As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
''Calc_MetrologyHSCnTS_ADV_Coefficient(
'hsc_nts_tg00i_temp_bts_ft1,
'BTS__Tg001_M1_TEMP10,
'hsc_nts_tg00i_d_0_ft2,
'hsc_nts_tg00i_d_0_FT1,
'hsc_nts_tp00g_c0a,
'hsc_nts_tp00g_c1a,
'hsc_nts_tp00g_c2a)

    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim str4 As String
    Dim str5 As String
    Dim str6 As String
    Dim str7 As String
    Dim str8 As String
        
    Dim site As Variant
    Dim i As Integer
    Dim x As New SiteLong
    Dim TestNameInput As String
    
    Dim Dict_BTS_CP1 As String
    Dim Dict_BTS_CP2 As String
    Dim Dict_12B32X As String
    Dim Dict_nTS_CP1 As String
    
    Dim Dict_BTS_CP1_Dec As New DSPWave
    Dim Dict_BTS_CP2_Dec As New DSPWave
    Dim Dict_12B32X_Dec As New DSPWave
    Dim Dict_nTS_CP1_Dec As New DSPWave
    
    Dim Dict_BTS_CP1_Bin As New DSPWave
    Dim Dict_BTS_CP2_Bin As New DSPWave
    Dim Dict_12B32X_Bin As New DSPWave
    Dim Dict_nTS_CP1_Bin As New DSPWave
        
    Dict_BTS_CP1 = argv(0)
    Dict_BTS_CP2 = argv(1)
    Dict_12B32X = argv(2)
    Dict_nTS_CP1 = argv(3)
    
    Dict_BTS_CP1_Bin = GetStoreDataAllType(Dict_BTS_CP1)
    Dict_BTS_CP2_Bin = GetStoreDataAllType(Dict_BTS_CP2)
    Dict_12B32X_Bin = GetStoreDataAllType(Dict_12B32X)
    Dict_nTS_CP1_Bin = GetStoreDataAllType(Dict_nTS_CP1)

    Call rundsp.BinToDec(Dict_BTS_CP1_Bin, Dict_BTS_CP1_Dec)
    Call rundsp.BinToDec(Dict_BTS_CP2_Bin, Dict_BTS_CP2_Dec)
    
    '20240422 Hidra by YM V02A new Coeff
    Dim Size_Dict_12B32X_Bin As New SiteLong
    For Each site In TheExec.sites
        Size_Dict_12B32X_Bin = Dict_12B32X_Bin.SampleSize
    Next site
    
    'Call rundsp.BinToDec(Dict_12B32X_Bin, Dict_12B32X_Dec)
    Call rundsp.DSP_2S_Complement_To_SignDec(Dict_12B32X_Bin, Size_Dict_12B32X_Bin, Dict_12B32X_Dec)
    '20240422 Hidra by YM V02A new Coeff
    
    Call rundsp.BinToDec(Dict_nTS_CP1_Bin, Dict_nTS_CP1_Dec)


    '------------------
'
    Dim Temp_tp00_c1a As New DSPWave
    Dim Temp_tp00_c1a_raw As New DSPWave
    Dim Temp_tp00_c1a_m As New DSPWave
    '20240422 Hidra by YM V02A new Coeff
    Dim Temp_tp00_c1a_term As New DSPWave
    Dim Temp_tp00_c1a_scaled As New DSPWave
    Dim Temp_tp00_c1a_K As New DSPWave
    Dim FT2_temp As Double: FT2_temp = 85
    '20240422 Hidra by YM V02A new Coeff
    Dim Temp_tp00_c1a_2sComplement As New DSPWave
    Dim Temp_tp00_c1a_2sComplement_Bin As New DSPWave
    Dim hsc_nts_tp00_c1a As String
    Dim Width As New SiteLong
    Width = 21
    ' Wait check golden
    'Dim a2_gold As Double: a2_gold = -141.26
    
    '====================20240814 Hidra B0 by YM====================
    Dim BKM_DECODE As String
    Dim sBKM As String
    
    For Each site In TheExec.sites
        BKM_DECODE = gS_BKM_IEDA
        Exit For
    Next site
    
    Dim a0_gold_adv As Double
    Dim a1_gold_adv As Double
    Dim a2_gold_adv As Double
    
    If MTR_Coefficient_CLS.BKM(UCase("NTS_N3EP_" & BKM_DECODE)) Then
        a0_gold_adv = MTR_Coefficient_CLS.CoefficientRead(UCase("NTS_N3EP_" & BKM_DECODE & "_a0_gold_adv"))
        a1_gold_adv = MTR_Coefficient_CLS.CoefficientRead(UCase("NTS_N3EP_" & BKM_DECODE & "_a1_gold_adv"))
        a2_gold_adv = MTR_Coefficient_CLS.CoefficientRead(UCase("NTS_N3EP_" & BKM_DECODE & "_a2_gold_adv"))
        sBKM = MTR_Coefficient_CLS.CoefficientRead(UCase("NTS_N3EP_" & BKM_DECODE & "_BKM_Version"))
    Else
        a0_gold_adv = MTR_Coefficient_CLS.CoefficientRead(UCase("NTS_N3EP_Default_a0_gold_adv"))
        a1_gold_adv = MTR_Coefficient_CLS.CoefficientRead(UCase("NTS_N3EP_Default_a1_gold_adv"))
        a2_gold_adv = MTR_Coefficient_CLS.CoefficientRead(UCase("NTS_N3EP_Default_a2_gold_adv"))
        sBKM = MTR_Coefficient_CLS.CoefficientRead(UCase("NTS_N3EP_Default_BKM_Version"))
    End If
    
    TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM : " & sBKM & " ) **********************"
    TheExec.Datalog.WriteComment "  a0_gold_adv = " & a0_gold_adv
    TheExec.Datalog.WriteComment "  a1_gold_adv = " & a1_gold_adv
    TheExec.Datalog.WriteComment "  a2_gold_adv = " & a2_gold_adv
    TheExec.Datalog.WriteComment "**************************************************************************"
    '====================20240814 Hidra B0 by YM====================
    
    hsc_nts_tp00_c1a = argv(5)
    Temp_tp00_c1a.CreateConstant 0, 1, DspDouble
    Temp_tp00_c1a_m.CreateConstant 0, 1, DspDouble
    Temp_tp00_c1a_raw.CreateConstant 0, 1, DspDouble
    '20240422 Hidra by YM V02A new Coeff
    Temp_tp00_c1a_term.CreateConstant 0, 1, DspDouble
    Temp_tp00_c1a_scaled.CreateConstant 0, 1, DspDouble
    Temp_tp00_c1a_K.CreateConstant 0, 1, DspDouble
    '20240422 Hidra by YM V02A new Coeff
    Temp_tp00_c1a_2sComplement_Bin.CreateConstant 0, 21, DspDouble  '20240131 Hidra by YM Fuse 21Bit
    For Each site In TheExec.sites.Active
        Temp_tp00_c1a = Dict_BTS_CP2_Dec.Subtract(Dict_BTS_CP1_Dec)     'hsc_nts_tg00i_temp_bts_FT2 - hsc_nts_tg00i_temp_bts_FT1
        Temp_tp00_c1a_m = Dict_12B32X_Dec.Subtract(Dict_nTS_CP1_Dec)    'hsc_nts_tg00i_d_0_FT2 - hsc_nts_tg00i_d_0_FT1
        Temp_tp00_c1a = Temp_tp00_c1a.divide(Temp_tp00_c1a_m)           '(hsc_nts_tg00i_temp_bts_FT2 - hsc_nts_tg00i_temp_bts_FT1) / (hsc_nts_tg00i_d_0_FT2 - hsc_nts_tg00i_d_0_FT1)
        Temp_tp00_c1a_raw = Temp_tp00_c1a.Multiply(2 ^ 10)             '(2^16 / 2^6) * (hsc_nts_tg00i_temp_bts_FT2 - hsc_nts_tg00i_temp_bts_FT1) / (hsc_nts_tg00i_d_0_FT2 - hsc_nts_tg00i_d_0_FT1)
        
        '20240422 Hidra by YM V02A new Coeff
        Temp_tp00_c1a_K = Temp_tp00_c1a_raw.Pow(-2).Multiply(a2_gold_adv).Multiply(FT2_temp - 25).Negate.Add(1)     'K = [1 - (FT2_temp - 25C) * a2_gold_adv * 1 / (Temp_tp00_c1a_raw ^ 2)
        Temp_tp00_c1a_scaled = Temp_tp00_c1a_K.Multiply(Temp_tp00_c1a_raw)      'Temp_tp00_c1a_scaled = K * Temp_tp00_c1a_raw
        'Hidra new 1st term
        Temp_tp00_c1a_term = Dict_nTS_CP1_Dec.divide(2 ^ 16).Multiply(2 * a2_gold_adv)      '2 * a2_gold_adv * (1 / 2^16) * (hsc_nts_tg00i_d_0_FT1)
        
        Temp_tp00_c1a = Temp_tp00_c1a_scaled.Subtract(Temp_tp00_c1a_term)       'Temp_tp00_c1a_scaled - 2 * a2_gold_adv * (1 / 2^16) * (hsc_nts_tg00i_d_0_FT1)
        
        Temp_tp00_c1a = Temp_tp00_c1a.Multiply(2 ^ 9)      ' 2^9 * [Temp_tp00_c1a_scaled - 2 * a2_gold_adv * (1 / 2^16) * (hsc_nts_tg00i_d_0_FT1)]
        '20240422 Hidra by YM V02A new Coeff
    Next site
    Call rundsp.DSP_Convert_2S_Complement(Temp_tp00_c1a, Width, Temp_tp00_c1a_2sComplement)
    Call rundsp.DSPWaveDecToBinary(Temp_tp00_c1a_2sComplement, 21, Temp_tp00_c1a_2sComplement_Bin)  '20240131 Hidra by YM Fuse 21Bit
    Call StoreDataAllType(hsc_nts_tp00_c1a, Temp_tp00_c1a_2sComplement_Bin)
    
    Dim Temp_tp00_c0a As New DSPWave
    Dim Temp_tp00_c0a_m As New DSPWave
    Dim Temp_tp00_c0a_new As New DSPWave
    Dim Temp_tp00_c0a_2sComplement As New DSPWave
    Dim Temp_tp00_c0a_2sComplement_Bin As New DSPWave
    Dim hsc_nts_tp00_c0a As String
    
    hsc_nts_tp00_c0a = argv(4)
    Temp_tp00_c0a.CreateConstant 0, 1, DspDouble
    Temp_tp00_c0a_new.CreateConstant 0, 1, DspDouble
    Temp_tp00_c0a_2sComplement_Bin.CreateConstant 0, 21, DspDouble  '20240131 Hidra by YM Fuse 21Bit
    
    For Each site In TheExec.sites.Active
        Temp_tp00_c0a = Dict_BTS_CP1_Dec.divide(2 ^ 6)      '(1 / 2^6) * hsc_nts_tg00i_temp_bts_FT1
        
        '(1 / 2^6) * hsc_nts_tg00i_temp_bts_FT1 - (1 / 2^16) * Temp_tp00_c1a_scaled * hsc_nts_tg00i_d_0_FT1
        Temp_tp00_c0a = Temp_tp00_c0a.Subtract(Temp_tp00_c1a_scaled.Multiply(Dict_nTS_CP1_Dec).divide(2 ^ 16))
        
        'Hidra new 2nd term
        Temp_tp00_c0a_new = Dict_nTS_CP1_Dec.divide(2 ^ 16).Pow(2).Multiply(a2_gold_adv).Add(a0_gold_adv)    'a2_gold_adv * [(1 / 2^16) * (hsc_nts_tg00i_d_0_FT1)]^2 + a0_gold_adv
        
        '(1 / 2^6) * hsc_nts_tg00i_temp_bts_FT1 - (1 / 2^16) * Temp_tp00_c1a_scaled * hsc_nts_tg00i_d_0_FT1 + a2_gold_adv * [(1 / 2^16) * (hsc_nts_tg00i_d_0_FT1)]^2 + a0_gold_adv
        Temp_tp00_c0a = Temp_tp00_c0a.Add(Temp_tp00_c0a_new)
        
        '2^9 * [(1 / 2^6) * hsc_nts_tg00i_temp_bts_FT1 - (1 / 2^16) * Temp_tp00_c1a_scaled * hsc_nts_tg00i_d_0_FT1 + a2_gold_adv * [(1 / 2^16) * (hsc_nts_tg00i_d_0_FT1)]^2 + a0_gold_adv]
        Temp_tp00_c0a = Temp_tp00_c0a.Multiply(2 ^ 9)
    Next site
    Call rundsp.DSP_Convert_2S_Complement(Temp_tp00_c0a, Width, Temp_tp00_c0a_2sComplement)
    Call rundsp.DSPWaveDecToBinary(Temp_tp00_c0a_2sComplement, 21, Temp_tp00_c0a_2sComplement_Bin)  '20240131 Hidra by YM Fuse 21Bit
    Call StoreDataAllType(hsc_nts_tp00_c0a, Temp_tp00_c0a_2sComplement_Bin)
    
    '20240131 Hidra by YM Fuse 21Bit
    Dim Temp_tp00_c2a As New DSPWave
    Dim hsc_nts_tp00_c2a As String
    Dim Temp_tp00_c2a_2sComplement As New DSPWave
    Dim Temp_tp00_c2a_2sComplement_Bin As New DSPWave
    
    hsc_nts_tp00_c2a = argv(6)
    Temp_tp00_c2a.CreateConstant (a2_gold_adv * 2 ^ 9), 1, DspDouble    '2^9 * a2_gold_adv
    Temp_tp00_c2a_2sComplement_Bin.CreateConstant 0, 21, DspDouble

    Call rundsp.DSP_Convert_2S_Complement(Temp_tp00_c2a, Width, Temp_tp00_c2a_2sComplement)
    Call rundsp.DSPWaveDecToBinary(Temp_tp00_c2a_2sComplement, 21, Temp_tp00_c2a_2sComplement_Bin)
    Call StoreDataAllType(hsc_nts_tp00_c2a, Temp_tp00_c2a_2sComplement_Bin)
    '20240131 Hidra by YM Fuse 21Bit
    
    For Each site In TheExec.sites.Active
        TheExec.Datalog.WriteComment "site_" & site & "_" & argv(0) & " : " & Dict_BTS_CP1_Dec(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & argv(1) & " : " & Dict_BTS_CP2_Dec(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & argv(2) & " : " & Dict_12B32X_Dec(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & argv(3) & " : " & Dict_nTS_CP1_Dec(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & hsc_nts_tp00_c0a & " : " & Temp_tp00_c0a_2sComplement(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & hsc_nts_tp00_c1a & " : " & Temp_tp00_c1a_2sComplement(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & hsc_nts_tp00_c2a & " : " & Temp_tp00_c2a_2sComplement(site).Element(0)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=Temp_tp00_c0a_2sComplement.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=Temp_tp00_c1a_2sComplement.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=Temp_tp00_c2a_2sComplement.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyHSCnTS_ADV_Coefficient") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
    
End Function




'[20231025][T-Bra][Chuan] MTRBTS post calculate function TTR optimize
Public Function Calc_MetrologyBTS_Coefficient_10b_N3E_TTR(argc As Integer, argv() As String) As Long
        On Error GoTo errHandler 'Add ErrHandler 2023/05/29
        
    Dim DSP_MTRBTS_OUT As New DSPWave
    Dim ARRAY_MTRBTS_OUT_DEC() As New SiteDouble
    Dim v0 As New SiteDouble
    Dim v1 As New SiteDouble
    ReDim ARRAY_MTRBTS_OUT_DEC(0)
    Dim x0(0) As New SiteDouble
    Dim x0a(0) As New SiteDouble
    Dim GC0, GC1, GC2, GC3 As Double
    Dim C0_CAL(0) As New SiteDouble
    Dim A0_CAL(0) As New SiteDouble
    Dim B0_CAL(0) As New SiteDouble
    Dim B1_CAL(0) As New SiteDouble
    Dim ARRAY_C0_CAL_eFuse() As Double
    Dim ARRAY_C0_CAL_SRC() As Long: ReDim ARRAY_C0_CAL_SRC(24)
    Dim TempVal As Double
    Dim TestNameInput As String
    Dim i As Long
    Dim k As Long
    Dim Sensor_Num() As String
    Dim BKM_DECODE As String
    Dim sBKM As String      'Update for parsing MTR table -- 20230103
    '===================201209 BTS new constant===========================
    Dim a0, a1, a2, a3, Tcp, k0, k1 As Double
    Dim sensor_loop_count As Integer
    Dim sensor_data() As String
    Dim All_sensor_C0 As New DSPWave
    Dim C0_max_value As New SiteDouble
    Dim C0_min_value As New SiteDouble
    Dim C0_avg_value As New SiteDouble
    Dim C0_max_value_index_dsp As New DSPWave
    Dim C0_min_value_index_dsp As New DSPWave
    
    Dim C0_max2avg_value As New SiteDouble
    Dim C0_min2avg_value As New SiteDouble
    
    
    sensor_loop_count = argc
    All_sensor_C0.CreateConstant 0, argc

    For k = 0 To sensor_loop_count - 1
        '====new===
        sensor_data = Split(argv(k), "+")
        '==========
        'Sensor_Num = Split(argv(0), "_") 'old
        Sensor_Num = Split(sensor_data(0), "_") 'new
        
        For Each site In TheExec.sites
            BKM_DECODE = gS_BKM_IEDA
            Exit For
        Next site
        
        If MTR_COEFF_INFO_DICT.Exists("BTS_N3E_" & BKM_DECODE) Then
            a3 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_A3"))
            a2 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_A2"))
            a1 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_A1"))
            a0 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_A0"))
            k1 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_K1"))
            k0 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_K0"))
            Tcp = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_TCP"))
            sBKM = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_BKM"))
        Else
            a3 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_A3"))
            a2 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_A2"))
            a1 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_A1"))
            a0 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_A0"))
            k1 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_K1"))
            k0 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_K0"))
            Tcp = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_TCP"))
            sBKM = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_BKM"))
        End If
    
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM :" & sBKM & ")**********************"
        TheExec.Datalog.WriteComment "  a3 = " & a3
        TheExec.Datalog.WriteComment "  a2 = " & a2
        TheExec.Datalog.WriteComment "  a1 = " & a1
        TheExec.Datalog.WriteComment "  a0 = " & a0
        TheExec.Datalog.WriteComment "  k1 = " & k1
        TheExec.Datalog.WriteComment "  k0 = " & k0
        TheExec.Datalog.WriteComment "  Tcp = " & Tcp
        TheExec.Datalog.WriteComment "*************************************************************************************"
    
    
        If gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "a0=" & a0 & ",a1=" & a1 & ",a2=" & a2 & ",a3=" & a3 & ",k0=" & k0 & ",k1=" & k1 & ",Tcp=" & Tcp
        End If
        ''===================================================================
        'ARRAY_MTRBTS_OUT_DEC(0) = GetStoredData(argv(0) & "_para") 'old
        ARRAY_MTRBTS_OUT_DEC(0) = GetStoredData(sensor_data(0) & "_" & sensor_data(1) & "_para") 'new
        'v0 = GetStoredData(argv(1)) 'old
        v0 = GetStoredData(sensor_data(0) & "_" & sensor_data(2)) 'new
        'v1 = GetStoredData(argv(2)) 'old
        v1 = GetStoredData(sensor_data(0) & "_" & sensor_data(3)) 'new
        For Each site In TheExec.sites.Active
            If v1 = 0 Then
                v1 = 0.00000001
                TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Coefficient V1 value is zero!!! "
            End If
            x0a(0) = ARRAY_MTRBTS_OUT_DEC(0) / (2 ^ 16) ''201209 for new BTS rule
            x0(0) = (x0a(0) - v0) / v1
            A0_CAL(0) = Tcp - a1 * x0(0) - a2 * x0(0) ^ 2 - a3 * x0(0) ^ 3
            B1_CAL(0) = k0 + k1 * x0(0)
            B0_CAL(0) = Tcp - B1_CAL(0) * x0(0) - a2 * x0(0) ^ 2 - a3 * x0(0) ^ 3
        Next site
        '===================201209 BTS new rule===========================
        Dim C0 As New SiteDouble
        Dim C1 As New SiteDouble
        Dim C2 As New SiteDouble
        Dim C3 As New SiteDouble
        Dim C0a As New SiteDouble
        Dim C1a As New SiteDouble
        Dim DSP_C0_Cal_eFuse As New DSPWave
        Dim DSP_C1_Cal_eFuse As New DSPWave
        Dim DSP_C2_Cal_eFuse As New DSPWave
        Dim DSP_C3_Cal_eFuse As New DSPWave
        Dim DSP_C0a_Cal_eFuse As New DSPWave
        Dim DSP_C1a_Cal_eFuse As New DSPWave
        Dim DSP_C0_CAL  As New DSPWave
        Dim DSP_C1_CAL  As New DSPWave
        Dim DSP_C2_CAL  As New DSPWave
        Dim DSP_C3_CAL  As New DSPWave
        Dim DSP_C0a_CAL As New DSPWave
        Dim DSP_C1a_CAL As New DSPWave
        Dim DSP_C0_Cal_Src As New DSPWave
        Dim DSP_C1_Cal_Src As New DSPWave
        Dim DSP_C2_Cal_Src As New DSPWave
        Dim DSP_C3_Cal_Src As New DSPWave
        Dim DSP_C0a_Cal_Src As New DSPWave
        Dim DSP_C1a_Cal_Src As New DSPWave
        
        
        Set DSP_C0_Cal_eFuse = Nothing
        Set DSP_C1_Cal_eFuse = Nothing
        Set DSP_C2_Cal_eFuse = Nothing
        Set DSP_C3_Cal_eFuse = Nothing
        Set DSP_C0a_Cal_eFuse = Nothing
        Set DSP_C1a_Cal_eFuse = Nothing
        Set DSP_C0_CAL = Nothing
        Set DSP_C1_CAL = Nothing
        Set DSP_C2_CAL = Nothing
        Set DSP_C3_CAL = Nothing
        Set DSP_C0a_CAL = Nothing
        Set DSP_C1a_CAL = Nothing
        
        'If k = 0 Then
        DSP_C0_Cal_eFuse.CreateConstant 0, 1
        DSP_C1_Cal_eFuse.CreateConstant 0, 1
        DSP_C2_Cal_eFuse.CreateConstant 0, 1
        DSP_C3_Cal_eFuse.CreateConstant 0, 1
        DSP_C0a_Cal_eFuse.CreateConstant 0, 1
        DSP_C1a_Cal_eFuse.CreateConstant 0, 1
        DSP_C0_CAL.CreateConstant 0, 1
        DSP_C1_CAL.CreateConstant 0, 1
        DSP_C2_CAL.CreateConstant 0, 1
        DSP_C3_CAL.CreateConstant 0, 1
        DSP_C0a_CAL.CreateConstant 0, 1
        DSP_C1a_CAL.CreateConstant 0, 1
        'End If
        
        C0 = A0_CAL(0).Subtract(v0.Multiply(a1).divide(v1))                                    ''A0_CAL(0)- a1 * V0 / V1
        C1 = v1.Invert.Multiply(a1).Subtract(v0.Multiply(2 * a2).divide(v1.power(2)))          ''a1 / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
        C2 = v1.power(2).Invert.Multiply(a2).Subtract(v0.Multiply(3 * a3).divide(v1.power(3))) ''a2 / (BTS_V1 ^ 2) - 3 * a3 * BTS_V0 / (BTS_V1 ^ 3)
        C3 = v1.power(3).Invert.Multiply(a3)                                                   ''a3 / (BTS_V1 ^ 3)
        C0a = B0_CAL(0).Subtract(B1_CAL(0).Multiply(v0).divide(v1))                            ''B0_CAL(0) - B1_CAL(0) * BTS_V0 / BTS_V1
        C1a = B1_CAL(0).divide(v1).Subtract(v0.Multiply(2 * a2).divide(v1.power(2)))           ''B1_CAL / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
        '=================================================================
        For Each site In TheExec.sites.Active
            DSP_C0_CAL.Element(0) = C0
            DSP_C1_CAL.Element(0) = C1
            DSP_C2_CAL.Element(0) = C2
            DSP_C3_CAL.Element(0) = C3
            DSP_C0a_CAL.Element(0) = C0a
            DSP_C1a_CAL.Element(0) = C1a
            
            All_sensor_C0.Element(k) = C0
        Next site
        If gl_Disable_HIP_debug_log = False Then
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=x0a(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=x0(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=A0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=B0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=B1_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=DSP_C0_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=DSP_C1_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=DSP_C2_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=DSP_C3_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=DSP_C0a_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=DSP_C1a_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        Else
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 11
        End If
        
        C0 = C0.Multiply(2 ^ 9)
        C1 = C1.Multiply(2 ^ 9)
        C2 = C2.Multiply(2 ^ 9)
        C3 = C3.Multiply(2 ^ 9)
        C0a = C0a.Multiply(2 ^ 9)
        C1a = C1a.Multiply(2 ^ 9)
        For Each site In TheExec.sites.Active
            DSP_C0_CAL.Element(0) = C0
            DSP_C1_CAL.Element(0) = C1
            DSP_C2_CAL.Element(0) = C2
            DSP_C3_CAL.Element(0) = C3
            DSP_C0a_CAL.Element(0) = C0a
            DSP_C1a_CAL.Element(0) = C1a
        Next site
        
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0_Cal_eFuse, DSP_C0_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1_Cal_eFuse, DSP_C1_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C2_Cal_eFuse, DSP_C2_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C3_Cal_eFuse, DSP_C3_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0a_Cal_eFuse, DSP_C0a_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1a_Cal_eFuse, DSP_C1a_CAL, 20, 0)
        
        
        Set DSP_C0_Cal_Src = Nothing
        Set DSP_C1_Cal_Src = Nothing
        Set DSP_C2_Cal_Src = Nothing
        Set DSP_C3_Cal_Src = Nothing
        Set DSP_C0a_Cal_Src = Nothing
        Set DSP_C1a_Cal_Src = Nothing
        
        Call HardIP_Dec2Bin(DSP_C0_Cal_Src, DSP_C0_Cal_eFuse, 20) 'DSP_DecToBin
        Call HardIP_Dec2Bin(DSP_C1_Cal_Src, DSP_C1_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C2_Cal_Src, DSP_C2_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C3_Cal_Src, DSP_C3_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C0a_Cal_Src, DSP_C0a_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C1a_Cal_Src, DSP_C1a_Cal_eFuse, 20)
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C0=" & DSP_C0_CAL.Element(0) & " => " & DSP_C0_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C1=" & DSP_C1_CAL.Element(0) & " => " & DSP_C1_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C2=" & DSP_C2_CAL.Element(0) & " => " & DSP_C2_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C3=" & DSP_C3_CAL.Element(0) & " => " & DSP_C3_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C0A=" & DSP_C0a_CAL.Element(0) & " => " & DSP_C0a_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C1A=" & DSP_C1a_CAL.Element(0) & " => " & DSP_C1a_Cal_eFuse.Element(0)
            Next site
        End If
        
   
        Call AddStoredCaptureData(Sensor_Num(0) & "_C0_SRC", DSP_C0_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C1_SRC", DSP_C1_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C2_SRC", DSP_C2_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C3_SRC", DSP_C3_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C0A_SRC", DSP_C0a_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C1A_SRC", DSP_C1a_Cal_Src)
        
    Next k
    
    Dim temp_max_sensor_name As String
    Dim temp_min_sensor_name As String
    
    
    For Each site In TheExec.sites.Active
    
        C0_max_value = All_sensor_C0.CalcMaximumValue
        C0_avg_value = All_sensor_C0.CalcMean
        C0_min_value = All_sensor_C0.CalcMinimumValue
        C0_max2avg_value = C0_max_value - C0_avg_value
        C0_min2avg_value = C0_min_value - C0_avg_value


        C0_max_value_index_dsp = All_sensor_C0.FindIndices(EqualTo, C0_max_value)
        C0_min_value_index_dsp = All_sensor_C0.FindIndices(EqualTo, C0_min_value)
        
        If C0_max_value_index_dsp.SampleSize = 0 Then
            TheExec.Datalog.WriteComment "[ERROR] Did't find site" & site & "max Coefficient C0 !!!!"
            Exit For
        Else
            For i = 0 To C0_max_value_index_dsp.SampleSize - 1
                If i = 0 Then
                    temp_max_sensor_name = Split(Split(argv(C0_max_value_index_dsp(site).data(i)), "+")(0), "_")(0)
                Else
                    temp_max_sensor_name = temp_max_sensor_name & "&" & Split(Split(argv(C0_max_value_index_dsp(site).data(i)), "+")(0), "_")(0)
                End If
            Next i
        End If
        
        If C0_min_value_index_dsp.SampleSize = 0 Then
            TheExec.Datalog.WriteComment "[ERROR] Did't find site" & site & "min Coefficient C0!!!!"
            Exit For
        Else
            For i = 0 To C0_min_value_index_dsp.SampleSize - 1
                If i = 0 Then
                    temp_min_sensor_name = Split(Split(argv(C0_min_value_index_dsp(site).data(i)), "+")(0), "_")(0)
                Else
                    temp_min_sensor_name = temp_min_sensor_name & "&" & Split(Split(argv(C0_min_value_index_dsp(site).data(i)), "+")(0), "_")(0)
                End If
            Next i
        End If
        
        TestNameInput = Report_TName_From_Instance("CalcC", temp_max_sensor_name)
        TheExec.flow.TestLimit resultVal:=C0_max_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        TestNameInput = Report_TName_From_Instance("CalcC", temp_min_sensor_name)
        TheExec.flow.TestLimit resultVal:=C0_min_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        TestNameInput = Report_TName_From_Instance("CalcC", "")
        TheExec.flow.TestLimit resultVal:=C0_avg_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        
        If gl_Disable_HIP_debug_log = False Then
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=C0_max2avg_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=C0_min2avg_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 5
        Else
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 3
        End If
        
    Next site

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Coefficient_10b_N3E_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

'[20231025][T-Bra][Chuan] MTRBTS post calculate function TTR optimize
Public Function Calc_MetrologyBTS_Temperature_TTR(argc As Integer, argv() As String) As Long
        On Error GoTo errHandler 'Add ErrHandler 2023/05/29
        
    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature As New DSPWave
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    Dim Sensor_Num() As String
    Dim Max_sensor_temperature As New SiteDouble
    Dim Min_sensor_temperature As New SiteDouble
    Dim Avg_sensor_temperature As New SiteDouble
    Dim Max2Avg_temperature As New SiteDouble
    Dim Min2Avg_temperature As New SiteDouble
    Dim max_temperature_index As New SiteDouble
    Dim min_temperature_index As New SiteDouble
    
    DSP_Temperature.CreateConstant 0, UBound(Temperature_Sensor) + 1

    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoredData(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature = Temperature.divide(64)
            DSP_Temperature.Element(i) = Temperature
        Next site
    Next i
    
    If gl_Disable_HIP_debug_log = False Then
        For i = 0 To UBound(Temperature_Sensor)
            Sensor_Num = Split(Temperature_Sensor(i), "_")
            TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
            TheExec.flow.TestLimit resultVal:=DSP_Temperature.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next i
    Else
        TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + UBound(Temperature_Sensor) + 1
    End If
    
    Dim temp_max_sensor_name As String
    Dim temp_min_sensor_name As String
    Dim max_temperature_index_dsp As New DSPWave
    Dim min_temperature_index_dsp As New DSPWave
    
    For Each site In TheExec.sites.Active
    
        
        Max_sensor_temperature = DSP_Temperature.CalcMaximumValue
        Min_sensor_temperature = DSP_Temperature.CalcMinimumValue
        Avg_sensor_temperature = DSP_Temperature.CalcMean
        Max2Avg_temperature = DSP_Temperature.CalcMaximumValue - DSP_Temperature.CalcMean
        Min2Avg_temperature = DSP_Temperature.CalcMinimumValue - DSP_Temperature.CalcMean

        max_temperature_index_dsp = DSP_Temperature.FindIndices(EqualTo, Max_sensor_temperature)
        min_temperature_index_dsp = DSP_Temperature.FindIndices(EqualTo, Min_sensor_temperature)
        
        If max_temperature_index_dsp.SampleSize = 0 Then
            TheExec.Datalog.WriteComment "[ERROR] Did't find site" & site & "max temperature!!!!"
            Exit For
        Else
            For i = 0 To max_temperature_index_dsp.SampleSize - 1
                If i = 0 Then
                    temp_max_sensor_name = Split(Temperature_Sensor(max_temperature_index_dsp(site).data(i)), "_")(0)
                Else
                    temp_max_sensor_name = temp_max_sensor_name & "&" & Split(Temperature_Sensor(max_temperature_index_dsp(site).data(i)), "_")(0)
                End If
            Next i
        End If
        
        If max_temperature_index_dsp.SampleSize = 0 Then
            TheExec.Datalog.WriteComment "[ERROR] Did't find site" & site & "min temperature!!!!"
            Exit For
        Else
            For i = 0 To min_temperature_index_dsp.SampleSize - 1
                If i = 0 Then
                    temp_min_sensor_name = Split(Temperature_Sensor(min_temperature_index_dsp(site).data(i)), "_")(0)
                Else
                    temp_min_sensor_name = temp_min_sensor_name & "&" & Split(Temperature_Sensor(min_temperature_index_dsp(site).data(i)), "_")(0)
                End If
            Next i
        End If

        
        TestNameInput = Report_TName_From_Instance("CalcC", temp_max_sensor_name)
        TheExec.flow.TestLimit resultVal:=Max_sensor_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        TestNameInput = Report_TName_From_Instance("CalcC", temp_min_sensor_name)
        TheExec.flow.TestLimit resultVal:=Min_sensor_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        TestNameInput = Report_TName_From_Instance("CalcC", "")
        TheExec.flow.TestLimit resultVal:=Avg_sensor_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        If gl_Disable_HIP_debug_log = False Then
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=Max2Avg_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.flow.TestLimit resultVal:=Min2Avg_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 5
        Else
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 3
        End If
        
    Next site
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Temperature_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
'20240131 Hidra by YM
Public Function Calc_MetrologyBTS_OffSet_TTR(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim InWf As New DSPWave
Dim InWf_SiteDouble As New SiteDouble
Dim Sensor_Store As String
Dim TestNameInput As String
Dim BTS_V0 As New SiteDouble
Dim i As Long
    For i = 0 To argc - 1
        'Sensor_Num = Split(argv(1), "_")
        Sensor_Store = Split(argv(i), "+")(1)
           
        'InWf_SiteDouble(0) = GetStoredData(argv(0) & "_para")
        InWf_SiteDouble = GetStoreDataAllType(Split(argv(i), "+")(0) & "_para")
        
        ''' Update for Donan @William 230110
        'BTS_V0 = InWf_SiteDouble(0).divide(2 ^ 16)
        BTS_V0 = InWf_SiteDouble.divide(2 ^ 16)
        
    '    BTS_V0 = InWf_SiteDouble(0).Subtract(2 ^ 16).divide(2 ^ 16) '.Abs.Multiply(0) 'V05A pipe
        TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
        TheExec.flow.TestLimit resultVal:=BTS_V0, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling ', formatStr:="%.5f"   ''new
           
        'Call AddStoredData(argv(1), BTS_V0)  '& "_" & CStr(TheExec.Flow.var(argv(2)).Value)
        Call StoreDataAllType(Sensor_Store, BTS_V0)
        
        Set InWf_SiteDouble = Nothing
        Set BTS_V0 = Nothing
        
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_OffSet_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function




'20240131 Hidra by YM
Public Function Calc_MetrologyBTS_Coefficient_TTR(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim DSP_MTRBTS_OUT As New DSPWave
    Dim ARRAY_MTRBTS_OUT_DEC() As New SiteDouble
    Dim v0 As New SiteDouble
    Dim v1 As New SiteDouble
    ReDim ARRAY_MTRBTS_OUT_DEC(0)
    Dim x0(0) As New SiteDouble
    Dim x0a(0) As New SiteDouble
    Dim GC0, GC1, GC2, GC3 As Double
    Dim C0_CAL(0) As New SiteDouble
    Dim A0_CAL(0) As New SiteDouble
    Dim B0_CAL(0) As New SiteDouble
    Dim B1_CAL(0) As New SiteDouble
    Dim ARRAY_C0_CAL_eFuse() As Double
    Dim ARRAY_C0_CAL_SRC() As Long: ReDim ARRAY_C0_CAL_SRC(24)
    Dim TempVal As Double
    Dim TestNameInput As String
    Dim i As Long
    '20240131 Hidra by YM
    Dim k As Long
    Dim Sensor_Num() As String
    'Sensor_Num = Split(argv(0), "_")
    ''' Update for Donan @William 220110
    Dim Sensor_Num_Str As String
    'Sensor_Num_Str = Sensor_Num(2)  ' Assign Sensor String
    '===================201209 BTS new constant===========================
    Dim a0, a1, a2, a3, Tcp, k0, k1 As Double
    '20240131 Hidra by YM
    Dim sensor_loop_count As Integer
    Dim sensor_data() As String
        
    'V1
'    A0 = 53.42539937
'    a1 = 106.24914282
'    a2 = -5.64960584
'    a3 = 0.5126488
'    k0 = 133.35292861
'    k1 = 98.97608898
'    Tcp = 24.3
    

    '20210510
'    A0 = 53.26973036
'    a1 = 104.64623585
'    a2 = -8.85459385
'    a3 = -0.30123576
'    k0 = 131.34112566
'    k1 = 97.48290549
'    Tcp = 25
    
    '20220609 Ibiza V05A V3 Walker
'    A0 = 64.78880945
'    a1 = 120.88266867
'    a2 = -8.65990009
'    a3 = -3.00683063
'    k0 = 120.88266867
'    k1 = 0
'    Tcp = 25
    
    '20231030 Donan V13A N3E-V2 William
    a0 = 65
    a1 = 120.0515752
    a2 = -5.30965391
    a3 = -0.26804929
    k0 = 120.0515752
    k1 = 0
    Tcp = 25

    
    '20240131 Hidra by YM
    sensor_loop_count = argc

    For k = 0 To sensor_loop_count - 1
        sensor_data = Split(argv(k), "+")
        
        'Sensor_Num = Split(argv(0), "_")
        Sensor_Num = Split(sensor_data(0), "_")
        Sensor_Num_Str = Sensor_Num(2)
    '20240131 Hidra by YM
    
        If gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "a0=" & a0 & ",a1=" & a1 & ",a2=" & a2 & ",a3=" & a3 & ",k0=" & k0 & ",k1=" & k1 & ",Tcp=" & Tcp
        End If
        ''===================================================================
        'ARRAY_MTRBTS_OUT_DEC(0) = GetStoredData(argv(0) & "_para")
        ARRAY_MTRBTS_OUT_DEC(0) = GetStoreDataAllType(sensor_data(0) & "_" & sensor_data(1) & "_para") 'new
        'v0 = GetStoredData(argv(1))
        v0 = GetStoreDataAllType(sensor_data(0) & "_" & sensor_data(2)) 'new
        'v1 = GetStoredData(argv(2))
        v1 = GetStoreDataAllType(sensor_data(0) & "_" & sensor_data(3)) 'new
        
        For Each site In TheExec.sites.Active
            If v1 = 0 Then
                v1 = 0.00000001
                TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Coefficient V1 value is zero!!! "
            End If
            x0a(0) = ARRAY_MTRBTS_OUT_DEC(0) / (2 ^ 16) ''201209 for new BTS rule
            x0(0) = (x0a(0) - v0) / v1
            A0_CAL(0) = Tcp - a1 * x0(0) - a2 * x0(0) ^ 2 - a3 * x0(0) ^ 3
            B1_CAL(0) = k0 + k1 * x0(0)
            B0_CAL(0) = Tcp - B1_CAL(0) * x0(0) - a2 * x0(0) ^ 2 - a3 * x0(0) ^ 3
        Next site
        '===================201209 BTS new rule===========================
        Dim C0 As New SiteDouble
        Dim C1 As New SiteDouble
        Dim C2 As New SiteDouble
        Dim C3 As New SiteDouble
        Dim C0a As New SiteDouble
        Dim C1a As New SiteDouble
        Dim DSP_C0_Cal_eFuse As New DSPWave
        Dim DSP_C1_Cal_eFuse As New DSPWave
        Dim DSP_C2_Cal_eFuse As New DSPWave
        Dim DSP_C3_Cal_eFuse As New DSPWave
        Dim DSP_C0a_Cal_eFuse As New DSPWave
        Dim DSP_C1a_Cal_eFuse As New DSPWave
        Dim DSP_C0_CAL  As New DSPWave
        Dim DSP_C1_CAL  As New DSPWave
        Dim DSP_C2_CAL  As New DSPWave
        Dim DSP_C3_CAL  As New DSPWave
        Dim DSP_C0a_CAL As New DSPWave
        Dim DSP_C1a_CAL As New DSPWave
        Dim DSP_C0_Cal_Src As New DSPWave
        Dim DSP_C1_Cal_Src As New DSPWave
        Dim DSP_C2_Cal_Src As New DSPWave
        Dim DSP_C3_Cal_Src As New DSPWave
        Dim DSP_C0a_Cal_Src As New DSPWave
        Dim DSP_C1a_Cal_Src As New DSPWave
        
        Dim AVG_C0_Cal_Src As New DSPWave
        
        Set DSP_C0_Cal_eFuse = Nothing
        Set DSP_C1_Cal_eFuse = Nothing
        Set DSP_C2_Cal_eFuse = Nothing
        Set DSP_C3_Cal_eFuse = Nothing
        Set DSP_C0a_Cal_eFuse = Nothing
        Set DSP_C1a_Cal_eFuse = Nothing
        Set DSP_C0_CAL = Nothing
        Set DSP_C1_CAL = Nothing
        Set DSP_C2_CAL = Nothing
        Set DSP_C3_CAL = Nothing
        Set DSP_C0a_CAL = Nothing
        Set DSP_C1a_CAL = Nothing
        
        Set AVG_C0_Cal_Src = Nothing
        
        DSP_C0_Cal_eFuse.CreateConstant 0, 1
        DSP_C1_Cal_eFuse.CreateConstant 0, 1
        DSP_C2_Cal_eFuse.CreateConstant 0, 1
        DSP_C3_Cal_eFuse.CreateConstant 0, 1
        DSP_C0a_Cal_eFuse.CreateConstant 0, 1
        DSP_C1a_Cal_eFuse.CreateConstant 0, 1
        DSP_C0_CAL.CreateConstant 0, 1
        DSP_C1_CAL.CreateConstant 0, 1
        DSP_C2_CAL.CreateConstant 0, 1
        DSP_C3_CAL.CreateConstant 0, 1
        DSP_C0a_CAL.CreateConstant 0, 1
        DSP_C1a_CAL.CreateConstant 0, 1
        
        AVG_C0_Cal_Src.CreateConstant 0, 1
    
        
        C0 = A0_CAL(0).Subtract(v0.Multiply(a1).divide(v1))                                    ''A0_CAL(0)- a1 * V0 / V1
        C1 = v1.Invert.Multiply(a1).Subtract(v0.Multiply(2 * a2).divide(v1.power(2)))          ''a1 / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
        C2 = v1.power(2).Invert.Multiply(a2).Subtract(v0.Multiply(3 * a3).divide(v1.power(3))) ''a2 / (BTS_V1 ^ 2) - 3 * a3 * BTS_V0 / (BTS_V1 ^ 3)
        C3 = v1.power(3).Invert.Multiply(a3)                                                   ''a3 / (BTS_V1 ^ 3)
        C0a = B0_CAL(0).Subtract(B1_CAL(0).Multiply(v0).divide(v1))                            ''B0_CAL(0) - B1_CAL(0) * BTS_V0 / BTS_V1
        C1a = B1_CAL(0).divide(v1).Subtract(v0.Multiply(2 * a2).divide(v1.power(2)))           ''B1_CAL / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
        '=================================================================
        For Each site In TheExec.sites.Active
            DSP_C0_CAL.Element(0) = C0
            DSP_C1_CAL.Element(0) = C1
            DSP_C2_CAL.Element(0) = C2
            DSP_C3_CAL.Element(0) = C3
            DSP_C0a_CAL.Element(0) = C0a
            DSP_C1a_CAL.Element(0) = C1a
            
            AVG_C0_Cal_Src.Element(0) = C0
            
        Next site
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=x0a(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=x0(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=A0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=B0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=B1_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_C0_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    Call AddStoredCaptureData(Sensor_Num_Str & "_C0_SRC", DSP_C0_Cal_Src)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_C1_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    Call AddStoredCaptureData(Sensor_Num_Str & "_C1_SRC", DSP_C1_Cal_Src)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_C2_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    Call AddStoredCaptureData(Sensor_Num_Str & "_C2_SRC", DSP_C2_Cal_Src)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_C3_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    Call AddStoredCaptureData(Sensor_Num_Str & "_C3_SRC", DSP_C3_Cal_Src)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_C0a_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    Call AddStoredCaptureData(Sensor_Num_Str & "_C0A_SRC", DSP_C0A_Cal_Src)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_C1a_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    Call AddStoredCaptureData(Sensor_Num_Str & "_C1A_SRC", DSP_C1A_Cal_Src)
    
        ''' Update for Donan @William 230906
        Dim sd_Stored_C0 As New SiteDouble
        For Each site In TheExec.sites
            sd_Stored_C0(site) = DSP_C0_CAL.Element(0)
        Next site
        Call StoreDataAllType(Sensor_Num_Str & "_C0_CAL", sd_Stored_C0)
        Set sd_Stored_C0 = Nothing
    '    GetStoredData (Sensor_Num_Str & "_C0_CAL")
        
        C0 = C0.Multiply(2 ^ 9)
        C1 = C1.Multiply(2 ^ 9)
        C2 = C2.Multiply(2 ^ 9)
        C3 = C3.Multiply(2 ^ 9)
        C0a = C0a.Multiply(2 ^ 9)
        C1a = C1a.Multiply(2 ^ 9)
        For Each site In TheExec.sites.Active
            DSP_C0_CAL.Element(0) = C0
            DSP_C1_CAL.Element(0) = C1
            DSP_C2_CAL.Element(0) = C2
            DSP_C3_CAL.Element(0) = C3
            DSP_C0a_CAL.Element(0) = C0a
            DSP_C1a_CAL.Element(0) = C1a
        Next site
        
    ''    For Each site In TheExec.sites.Active
    ''        DSP_C0_CAL.Element(0) = 0 'C0
    ''        DSP_C1_CAL.Element(0) = 0 'C0 '1
    ''        DSP_C2_CAL.Element(0) = 0 'C0 '2
    ''        DSP_C3_CAL.Element(0) = 0 'C3
    ''        DSP_C0A_CAL.Element(0) = 0 ' C0A
    ''        DSP_C1A_CAL.Element(0) = 0 ' C1A
    ''    Next site
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0_Cal_eFuse, DSP_C0_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1_Cal_eFuse, DSP_C1_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C2_Cal_eFuse, DSP_C2_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C3_Cal_eFuse, DSP_C3_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0a_Cal_eFuse, DSP_C0a_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1a_Cal_eFuse, DSP_C1a_CAL, 20, 0)
        
        Set DSP_C0_Cal_Src = Nothing
        Set DSP_C1_Cal_Src = Nothing
        Set DSP_C2_Cal_Src = Nothing
        Set DSP_C3_Cal_Src = Nothing
        Set DSP_C0a_Cal_Src = Nothing
        Set DSP_C1a_Cal_Src = Nothing
        
        Call HardIP_Dec2Bin(DSP_C0_Cal_Src, DSP_C0_Cal_eFuse, 20) 'DSP_DecToBin
        Call HardIP_Dec2Bin(DSP_C1_Cal_Src, DSP_C1_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C2_Cal_Src, DSP_C2_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C3_Cal_Src, DSP_C3_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C0a_Cal_Src, DSP_C0a_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C1a_Cal_Src, DSP_C1a_Cal_eFuse, 20)
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C0=" & DSP_C0_CAL.Element(0) & " => " & DSP_C0_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C1=" & DSP_C1_CAL.Element(0) & " => " & DSP_C1_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C2=" & DSP_C2_CAL.Element(0) & " => " & DSP_C2_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C3=" & DSP_C3_CAL.Element(0) & " => " & DSP_C3_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C0A=" & DSP_C0a_CAL.Element(0) & " => " & DSP_C0a_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C1A=" & DSP_C1a_CAL.Element(0) & " => " & DSP_C1a_Cal_eFuse.Element(0)
            Next site
        End If
        
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=x0a(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=x0(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=A0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=B0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=B1_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=DSP_C0_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        Call StoreDataAllType(Sensor_Num_Str & "_C0_SRC", DSP_C0_Cal_Src)
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=DSP_C1_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        Call StoreDataAllType(Sensor_Num_Str & "_C1_SRC", DSP_C1_Cal_Src)
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=DSP_C2_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        Call StoreDataAllType(Sensor_Num_Str & "_C2_SRC", DSP_C2_Cal_Src)
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=DSP_C3_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        Call StoreDataAllType(Sensor_Num_Str & "_C3_SRC", DSP_C3_Cal_Src)
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=DSP_C0A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        Call StoreDataAllType(Sensor_Num_Str & "_C0A_SRC", DSP_C0a_Cal_Src)
    '    TestNameInput = Report_TName_From_Instance("CalcC", "")
    '    TheExec.Flow.TestLimit resultVal:=DSP_C1A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        Call StoreDataAllType(Sensor_Num_Str & "_C1A_SRC", DSP_C1a_Cal_Src)
        
        
        Call StoreDataAllType(Sensor_Num_Str & "_C0_CAL_AVG", AVG_C0_Cal_Src)
        
    Next k
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Coefficient_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_DCTS_Calibration(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'Alg::Calc_DCTS_Calibration(DCTS_LDO_0+DCTS_LDO_1+DCTS_LDO_2+DCTS_LDO_3+DCTS_LDO_4+DCTS_LDO_5+DCTS_LDO_6+DCTS_LDO_7+DCTS_LDO_8+DCTS_LDO_9+DCTS_LDO_10+DCTS_LDO_11+DCTS_LDO_12+DCTS_LDO_13+DCTS_LDO_14+DCTS_LDO_15+DCTS_LDO_16+DCTS_LDO_17+DCTS_LDO_18+DCTS_LDO_19+DCTS_LDO_20+DCTS_LDO_21+DCTS_LDO_22+DCTS_LDO_23
',1
',pcie_st__ana_vdc_
',dcts_dc_tune_3b+dcts_ldo_sel_2b+dcts_config_3b
',3+2+3)
  
    Dim i, j As Long
    Dim TestNameInput As String
    Dim target As Double: target = CDbl(argv(1))
    Dim MeasAry() As String: MeasAry = Split(argv(0), "+")
    Dim DicAry() As String: DicAry = Split(argv(3), "+")
    Dim BestCalResult As New SiteDouble: BestCalResult = 999
    Dim TempCalResult As New SiteDouble
    Dim BestResult_Str As New SiteVariant
    Dim BinAry() As String
    Dim DSP_FinalResult() As New DSPWave: ReDim DSP_FinalResult(UBound(DicAry))
    Dim DSP_FinalResult_Dec() As New DSPWave: ReDim DSP_FinalResult_Dec(UBound(DicAry))
    Dim DSP_FinalResult_Bin() As New DSPWave: ReDim DSP_FinalResult_Bin(UBound(DicAry))
    Dim C_Width() As New SiteLong: ReDim C_Width(UBound(DicAry))
    
    For i = 0 To UBound(MeasAry)
        TempCalResult = GetStoreDataAllType(MeasAry(i))
        TempCalResult = TempCalResult.Subtract(target).Abs
        For Each site In TheExec.sites
            If TempCalResult < BestCalResult Then
                BestCalResult = TempCalResult
                BestResult_Str = Replace(LCase(Instance_Data.Tname(i)), LCase(argv(2)), "")
            End If
        Next site
    Next i
        
    For i = 0 To UBound(DicAry)
        For Each site In TheExec.sites
            BinAry = Split(BestResult_Str, "_")
            'Put assigment into DSPwave
            DSP_FinalResult(i).CreateConstant 0, Len(BinAry(i)), DspLong
            'pcie_st__ana_vdc_011_00_0 [Tname LSB is on the right side]
            For j = 0 To DSP_FinalResult(i).SampleSize - 1
                DSP_FinalResult(i).Element(j) = CLng(mid(BinAry(i), DSP_FinalResult(i).SampleSize - j, 1))  'pcie_st__ana_vdc_011_00_0 [Tname LSB is on the right side]
            Next j
        Next site
          
        'Binary to Dec
        DSP_FinalResult_Dec(i).CreateConstant 0, 1, DspLong
        rundsp.BinToDec DSP_FinalResult(i), DSP_FinalResult_Dec(i)
        
        'If Dic need Process
        If LCase(DicAry(i)) = LCase("dcts_config_3b") Then
            For Each site In TheExec.sites
                DSP_FinalResult_Dec(i).Element(0) = DSP_FinalResult_Dec(i).Element(0) * 4
            Next site
        End If
        
        'Dec to Binary
        C_Width(i) = Int(Split(argv(4), "+")(i))
        DSP_FinalResult_Bin(i).CreateConstant 0, C_Width(i), DspLong
        rundsp.DSPWaveDecToBinary DSP_FinalResult_Dec(i), C_Width(i), DSP_FinalResult_Bin(i)
        'Store
        Call StoreDataAllType(DicAry(i), DSP_FinalResult_Bin(i))
        
        TestNameInput = Report_TName_From_Instance(CalcV, "X", , , , , , , tlForceFlow)
        TheExec.flow.TestLimit resultVal:=DSP_FinalResult_Dec(i).Element(0), PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_DCTS_Calibration") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_DCTS_Calibration_UCSDM(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'Alg::Calc_DCTS_Calibration(DCTS_LDO_0+DCTS_LDO_1+DCTS_LDO_2+DCTS_LDO_3+DCTS_LDO_4+DCTS_LDO_5+DCTS_LDO_6+DCTS_LDO_7+DCTS_LDO_8+DCTS_LDO_9+DCTS_LDO_10+DCTS_LDO_11+DCTS_LDO_12+DCTS_LDO_13+DCTS_LDO_14+DCTS_LDO_15+DCTS_LDO_16+DCTS_LDO_17+DCTS_LDO_18+DCTS_LDO_19+DCTS_LDO_20+DCTS_LDO_21+DCTS_LDO_22+DCTS_LDO_23
',1
',pcie_st__ana_vdc_
',dcts_dc_tune_3b+dcts_ldo_sel_2b+dcts_config_3b
',3+2+3)
  
    Dim i, j As Long
    Dim TestNameInput As String
    Dim target As Double: target = CDbl(argv(1))
    Dim keyword As String: keyword = argv(2)
    Dim MeasAry() As String: MeasAry = Split(argv(0), "+")
    Dim DicAry() As String: DicAry = Split(argv(3), "+")
    Dim BestCalResult As New SiteDouble: BestCalResult = 999
    Dim TempCalResult As New SiteDouble
    Dim BestResult_Str As New SiteVariant
    Dim BinAry() As String
    Dim DSP_FinalResult() As New DSPWave: ReDim DSP_FinalResult(UBound(DicAry))
    Dim DSP_FinalResult_Dec() As New DSPWave: ReDim DSP_FinalResult_Dec(UBound(DicAry))
    Dim DSP_FinalResult_Bin() As New DSPWave: ReDim DSP_FinalResult_Bin(UBound(DicAry))
    Dim C_Width() As New SiteLong: ReDim C_Width(UBound(DicAry))
    
    For i = 0 To UBound(MeasAry)
        TempCalResult = GetStoreDataAllType(MeasAry(i) & "_para")
        TempCalResult = TempCalResult.Multiply(1.32).divide(4096)   'UCSDM -> Analog_Voltage = UCSDM_CODE x 1.32 / 4096
        TempCalResult = TempCalResult.Subtract(target).Abs
        For Each site In TheExec.sites
            If TempCalResult < BestCalResult Then
                BestCalResult = TempCalResult
                BestResult_Str = Replace(LCase(MeasAry(i)), LCase(keyword), "")
            End If
        Next site
    Next i
        
    For i = 0 To UBound(DicAry)
        For Each site In TheExec.sites
            BinAry = Split(BestResult_Str, "_")
            'Put assigment into DSPwave
            DSP_FinalResult(i).CreateConstant 0, Len(BinAry(i)), DspLong
            'pcie_st__ana_vdc_011_00_0 [Tname LSB is on the right side]
            For j = 0 To DSP_FinalResult(i).SampleSize - 1
                DSP_FinalResult(i).Element(j) = CLng(mid(BinAry(i), DSP_FinalResult(i).SampleSize - j, 1))  'pcie_st__ana_vdc_011_00_0 [Tname LSB is on the right side]
            Next j
        Next site
          
        'Binary to Dec
        DSP_FinalResult_Dec(i).CreateConstant 0, 1, DspLong
        rundsp.BinToDec DSP_FinalResult(i), DSP_FinalResult_Dec(i)
        
        'If Dic need Process
        If LCase(DicAry(i)) = LCase("dcts_config_3b") Then
            For Each site In TheExec.sites
                DSP_FinalResult_Dec(i).Element(0) = DSP_FinalResult_Dec(i).Element(0) * 4
            Next site
        End If
        
        'Dec to Binary
        C_Width(i) = Int(Split(argv(4), "+")(i))
        DSP_FinalResult_Bin(i).CreateConstant 0, C_Width(i), DspLong
        rundsp.DSPWaveDecToBinary DSP_FinalResult_Dec(i), C_Width(i), DSP_FinalResult_Bin(i)
        'Store
        Call StoreDataAllType(DicAry(i), DSP_FinalResult_Bin(i))
        
        TestNameInput = Report_TName_From_Instance(CalcV, "X", , , , , , , tlForceFlow)
        TheExec.flow.TestLimit resultVal:=DSP_FinalResult_Dec(i).Element(0), PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_DCTS_Calibration_UCSDM") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_DCTS_Coefficient(argc As Integer, argv() As String) As Long
    '20240131 Hidra by YM
    
    On Error GoTo errHandler
    
    Dim TestNameInput As String
    Dim input_value As New SiteDouble
    Dim C0_Value As New SiteDouble
    Dim storename As String
    Dim temp As String
    
    input_value = GetStoreDataAllType(argv(0) & "_para")
    input_value = input_value.divide(2 ^ 18)    'L
    
    If TheExec.TesterMode = testModeOffline Then
        input_value = 0.0005198
    End If
    
    'c0 = T - c1 * L - c2* L ^ 2 - c3 * L ^ 3 where T=25.0
    
    Dim C1, C2, C3, t As Double
    
    C1 = -197.3341
    C2 = -52.4061
    C3 = -111.8734
    t = 25
    
    TheExec.Datalog.WriteComment "********************** Coefficients for Calibration **********************"
    TheExec.Datalog.WriteComment "  c1 = " & C1
    TheExec.Datalog.WriteComment "  c2 = " & C2
    TheExec.Datalog.WriteComment "  c3 = " & C3
    TheExec.Datalog.WriteComment "  T  = " & t
    TheExec.Datalog.WriteComment "**************************************************************************"
    
    C0_Value = input_value.Multiply(C1).Negate.Add(t)
    C0_Value = C0_Value.Subtract(input_value.power(2).Multiply(C2))
    C0_Value = C0_Value.Subtract(input_value.power(3).Multiply(C3))
    
    TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=C0_Value, PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow

    Dim C0_Value_Dec As New DSPWave: C0_Value_Dec.CreateConstant 0, 1, DspDouble
    Dim C0_Value_2S_Dec As New DSPWave: C0_Value_2S_Dec.CreateConstant 0, 1, DspDouble
    Dim C0_width As New SiteLong: C0_width = 21
    Dim C0_Value_2S_Bin As New DSPWave: C0_Value_2S_Bin.CreateConstant 0, 21, DspLong
    
    For Each site In TheExec.sites
        C0_Value_Dec.Element(0) = C0_Value.Multiply(2 ^ 12)
    Next site
    
    Call rundsp.DSP_Convert_2S_Complement(C0_Value_Dec, C0_width, C0_Value_2S_Dec)
    Call rundsp.DSPWaveDecToBinary(C0_Value_2S_Dec, C0_width, C0_Value_2S_Bin)
    Call StoreDataAllType(argv(1), C0_Value_2S_Bin)


    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites.Active
            TheExec.Datalog.WriteComment "site_" & site & "_" & argv(1) & "_C0=" & C0_Value_Dec.Element(0) & " => " & C0_Value_2S_Dec.Element(0)
        Next site
    End If
        
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_DCTS_Coefficient") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_ICTS_Calibration(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
'Alg::Calc_ICTS_Calibration(
'ICTS__Te008_IREF_TRIM_MSB_1b+ICTS__Te008_IREF_TRIM_LSB_2b,
'8,
'1,
'ICTS_LDO,
'iref_trim_msb+iref_trim
'6+2)
    If ICTS_Sweep_Data_Store.Parse_Done = False Then
        ICTS_Sweep_Data_Store.index = 0
        ICTS_Sweep_Data_Store.LoopCount = CInt(argv(1)) - 1
        ICTS_Sweep_Data_Store.Meastore = argv(3)
        ReDim ICTS_Sweep_Data_Store.data(0)
        ICTS_Sweep_Data_Store.data(0) = GetStoreDataAllType(ICTS_Sweep_Data_Store.Meastore)
        ReDim ICTS_Sweep_Data_Store.Sweep_Ary(0)
        ICTS_Sweep_Data_Store.Sweep_Ary(0) = LCase(Instance_Data.DigSrc_Assignment)
        ICTS_Sweep_Data_Store.Parse_Done = True
        If ICTS_Sweep_Data_Store.index < ICTS_Sweep_Data_Store.LoopCount Then Exit Function
    Else
        ICTS_Sweep_Data_Store.index = ICTS_Sweep_Data_Store.index + 1
        ReDim Preserve ICTS_Sweep_Data_Store.data(ICTS_Sweep_Data_Store.index)
        ICTS_Sweep_Data_Store.data(ICTS_Sweep_Data_Store.index) = GetStoreDataAllType(ICTS_Sweep_Data_Store.Meastore)
        ReDim Preserve ICTS_Sweep_Data_Store.Sweep_Ary(ICTS_Sweep_Data_Store.index)
        ICTS_Sweep_Data_Store.Sweep_Ary(ICTS_Sweep_Data_Store.index) = LCase(Instance_Data.DigSrc_Assignment)
        If ICTS_Sweep_Data_Store.index < ICTS_Sweep_Data_Store.LoopCount Then Exit Function
    End If
    
    'For next instance
    ICTS_Sweep_Data_Store.Parse_Done = False
    
    Dim i, j, k As Long
    Dim TestNameInput As String
    Dim target As Double: target = CDbl(argv(2))
    Dim FinalResult As New SiteDouble: FinalResult = 999
    Dim TempResult As New SiteDouble
    Dim Final_Condition As New SiteVariant

            
    For i = 0 To ICTS_Sweep_Data_Store.LoopCount
        TempResult = ICTS_Sweep_Data_Store.data(i).pins(0).Subtract(target).Abs
        For Each site In TheExec.sites
            If TempResult < FinalResult Then
                FinalResult = TempResult
                Final_Condition = ICTS_Sweep_Data_Store.Sweep_Ary(i)
            End If
        Next site
    Next i
    
    Dim Dic_ary() As String: Dic_ary = Split(argv(0), "+")
    Dim Dic_Ary_temp() As String
    Dim Dic_temp() As String
   
    For i = 0 To UBound(Dic_ary)
    
        Dim C_Width As New SiteLong
        Dim DSP_FinalResult As New DSPWave
        Dim DSP_FinalResult_Bin As New DSPWave
        Dim DSP_FinalResult_Dec As New DSPWave
        
        For Each site In TheExec.sites
            Dic_Ary_temp = Split(Final_Condition, ";")
            For j = 0 To UBound(Dic_Ary_temp)
                If InStr(Dic_Ary_temp(j), LCase(Dic_ary(i))) > 0 Then
                    'Put assigment into DSPwave
                    Dic_temp = Split(Dic_Ary_temp(j), "=")
                    DSP_FinalResult.CreateConstant 0, Len(Dic_temp(1))
                    For k = 0 To DSP_FinalResult.SampleSize - 1
                        DSP_FinalResult.Element(k) = CLng(mid(Dic_temp(1), k + 1, 1))
                    Next k
                    Exit For
                End If
            Next j
        Next site
        'Binary to Dec
        DSP_FinalResult_Dec.CreateConstant 0, 1, DspLong
        rundsp.BinToDec DSP_FinalResult, DSP_FinalResult_Dec
        'If Dic need Process
        If LCase(Dic_ary(i)) = "icts__te008_iref_trim_msb_1b" Then
            For Each site In TheExec.sites
                DSP_FinalResult_Dec.Element(0) = DSP_FinalResult_Dec.Element(0) * 8
            Next site
        End If
        'Dec to Binary
        C_Width = Int(Split(argv(5), "+")(i))
        DSP_FinalResult_Bin.CreateConstant 0, Int(Split(argv(5), "+")(i)), DspLong
        rundsp.DSPWaveDecToBinary DSP_FinalResult_Dec, C_Width, DSP_FinalResult_Bin
        'Store Dictionary
        Call StoreDataAllType(Dic_ary(i), DSP_FinalResult_Bin)
        TestNameInput = Report_TName_From_Instance(CalcV, "X", , , , , , , tlForceFlow)
        TheExec.flow.TestLimit resultVal:=DSP_FinalResult_Dec.Element(0), PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
         
        Set DSP_FinalResult = Nothing
        Set DSP_FinalResult_Bin = Nothing
        Set DSP_FinalResult_Dec = Nothing
        Set C_Width = Nothing
    Next i

 Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_ICTS_Calibration") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_ICTS_RHIMCAL(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
    '---- New Function : Use for SEPF2M -- Updated to 3 argument for T-Har
    'Alg::Calc_ICTS_RHIMCAL(input1, input2, output, bit, forceI_value)                         -- Format
    'Alg::Calc_ICTS_RHIMCAL(ICTS_RHIM_INSTx_VSP,ICTS_RHIM_INSTx_VSN,OUTPUT_DICT_NAME,16,0.000010)               -- Example
    'Formula : R= (ICTS_RHIM_INSTx_VSP - ICTS_RHIM_INSTx_VSN) / 27e-6, unit is in Ohm

    Dim ICTS_RHIM_INSTx_VSP As New SiteDouble
    Dim ICTS_RHIM_INSTx_VSN As New SiteDouble
    Dim ICTS_RHIM_INSTx_VSP_ENG As New PinListData
    Dim ICTS_RHIM_INSTx_VSN_ENG As New PinListData
    
    Dim OUTPUT_DICT_NAME As String
    Dim Calc_R As New SiteDouble
    Dim DSPWF_Calc_R As New DSPWave
    DSPWF_Calc_R.CreateConstant 0, 1, DspDouble
    
    Dim DSPWF_BIT_SIZE As Long
    Dim DSPWF_R_BIN As New DSPWave
    Dim vsite As Variant
    Dim TestNameInput As String
    Dim ENG_String As String
    Dim k As Double
    Dim vest As Double
    Dim p1 As Double
    Dim forceI_value As Double: forceI_value = CDbl(argv(4))
    
'    If TheExec.TesterMode = testModeOffline Then
'        ICTS_RHIM_INSTx_VSP_ENG.AddPin ("dummy")
'        For Each vsite In TheExec.sites
'            ICTS_RHIM_INSTx_VSP_ENG.Pins("dummy").value = 0.5
'        Next vsite
'        ENG_String = argv(0)
'        Call AddStoredMeasurement(ENG_String, ICTS_RHIM_INSTx_VSP_ENG)
'
'        ICTS_RHIM_INSTx_VSN_ENG.AddPin ("dummy")
'        For Each vsite In TheExec.sites
'            ICTS_RHIM_INSTx_VSN_ENG.Pins("dummy").value = 0.4
'        Next vsite
'        ENG_String = argv(1)
'        Call AddStoredMeasurement(ENG_String, ICTS_RHIM_INSTx_VSN_ENG)
'    End If
    
    ICTS_RHIM_INSTx_VSP = GetStoreDataAllType(argv(0))
    ICTS_RHIM_INSTx_VSN = GetStoreDataAllType(argv(1))
    OUTPUT_DICT_NAME = argv(2)
    DSPWF_BIT_SIZE = CDbl(argv(3))
    
    'Formula : R= (ICTS_RHIM_INSTx_VSP - ICTS_RHIM_INSTx_VSN) / 27e-6, unit is in Ohm
    Calc_R = ICTS_RHIM_INSTx_VSP.Subtract(ICTS_RHIM_INSTx_VSN).divide(forceI_value).Add(0.5).Truncate
    
    For Each vsite In TheExec.sites
        DSPWF_Calc_R.Element(0) = Calc_R(vsite)
        TheExec.Datalog.WriteComment ("site : " & vsite & ", R = " & Calc_R(vsite) & " Ohm")
    Next vsite
    
    'Print out datalog
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=DSPWF_Calc_R.Element(0), PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow, unit:=unitCustom, customUnit:=" Ohm"
    
    Call HardIP_Dec2Bin(DSPWF_R_BIN, DSPWF_Calc_R, DSPWF_BIT_SIZE)
    Call StoreDataAllType(OUTPUT_DICT_NAME, DSPWF_R_BIN)
    
    'Store Double R value
    Call StoreDataAllType(OUTPUT_DICT_NAME & "_para", Calc_R)
   
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_ICTS_RHIMCAL") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MetrologyBTS_Temperature_T5P4(argc As Integer, argv() As String) As Long

    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim TemperatureTemp As Double
    Dim Temperature_Array_1() As New SiteDouble: ReDim Temperature_Array_1(UBound(Split(argv(0), "+")))
    Dim Temperature_Array_2() As New SiteDouble: ReDim Temperature_Array_2(UBound(Split(argv(0), "+")))
    Dim Temperature_Array_3() As New SiteDouble: ReDim Temperature_Array_3(UBound(Split(argv(0), "+")))
'    Dim Temperature_Array_SiteDouble() As New SiteDouble: ReDim Temperature_Array_SiteDouble(UBound(Split(argv(0), "+")))
    Dim VoltageValue As Double
    Dim DivValue As New SiteDouble
'    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
'    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
'    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    Dim i As Long
    Dim j As Long
    Dim site As Variant
    Dim TestNameInput As String
'    Dim Temperature_Dictionary() As String
'    Dim Sensor_Num() As String
    
'    Dim DSP_MTRSNS_Temperature_Average As New DSPWave
'    Dim DSP_MTRSNS_Temperature_Maximum As New DSPWave
'    Dim DSP_MTRSNS_Temperature_Minimum As New DSPWave
'    Dim DSP_MTRSNS_Temperature_Maximum_2 As New DSPWave
    
'    DSP_MTRSNS_Temperature_Average.CreateConstant 0, UBound(Temperature_Sensor) + 1
'    DSP_MTRSNS_Temperature_Maximum.CreateConstant 0, UBound(Temperature_Sensor) + 1
'    DSP_MTRSNS_Temperature_Minimum.CreateConstant 0, UBound(Temperature_Sensor) + 1
'    DSP_MTRSNS_Temperature_Maximum_2.CreateConstant 0, UBound(Temperature_Sensor) + 1
'    For Each Site In TheExec.sites
        For j = 1 To 3
            For i = 0 To UBound(Temperature_Sensor)
                Temperature = GetStoredData(Temperature_Sensor(i) & "_" & j & "_para")
                For Each site In TheExec.sites
                
'                    If Temperature >= 65536 Then
'                        TemperatureTemp = Abs((Temperature - 131072) / 131072)
'                    Else
'                        TemperatureTemp = Abs(Temperature / 131072)
'                    End If
                    
                    If Temperature >= 65536 Then
                        TemperatureTemp = ((Temperature - 131072) / 131072) * -1
                    Else
                        TemperatureTemp = (Temperature / 131072) * -1
                    End If
                    
                    
                    
                    If LCase(Temperature_Sensor(i) & "_" & j) Like "*_tp*" Or LCase(Temperature_Sensor(i) & "_" & j) Like "*_te*" Then
                        VoltageValue = TheHdw.DCVS.pins("VDD_FIXED").Voltage.value
                        TemperatureTemp = TemperatureTemp * VoltageValue
                    ElseIf LCase(Temperature_Sensor(i) & "_" & j) Like "*_ts*" Or LCase(Temperature_Sensor(i) & "_" & j) Like "*_tg*" Then
                        VoltageValue = TheHdw.DCVS.pins("VDD_SRAM_SOC").Voltage.value
                        TemperatureTemp = TemperatureTemp * VoltageValue
                    ElseIf LCase(Temperature_Sensor(i) & "_" & j) Like "*_ta*" Then
                        VoltageValue = TheHdw.DCVS.pins("VDD_SRAM_SOC").Voltage.value
                        TemperatureTemp = TemperatureTemp * VoltageValue
                    End If
                        
                    If j = 1 Then
                        Temperature_Array_1(i) = TemperatureTemp
                    ElseIf j = 2 Then
                        Temperature_Array_2(i) = TemperatureTemp
                    ElseIf j = 3 Then
                        Temperature_Array_3(i) = TemperatureTemp
                    End If
                Next site
            Next i
        Next j
'    Next Site
        For i = 0 To UBound(Temperature_Sensor)
                TestNameInput = Report_TName_From_Instance("CalcC", "")
                TheExec.flow.TestLimit resultVal:=Temperature_Array_1(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next i
        For j = 1 To 2
            For i = 0 To UBound(Temperature_Sensor)
            For Each site In TheExec.sites
                    If j = 1 Then
                            DivValue = Temperature_Array_2(i) - Temperature_Array_1(i)
                    ElseIf j = 2 Then
                            DivValue = Temperature_Array_3(i) - Temperature_Array_1(i)
                    End If
            Next site
                    TestNameInput = Report_TName_From_Instance("CalcC", "")
                    TheExec.flow.TestLimit resultVal:=DivValue, Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next i
        Next j
    
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Temperature"
    If AbortTest Then Exit Function Else Resume Next

End Function


' [20230908][All][Neil] MTRBTS DELTA CALCULATIONS TESTS
Public Function Calc_MetrologyBTS_Delta_MAX_MIN_AVG(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
    'Must to Reference "Calc_MetrologyBTS_Coefficient" function output result
    'Format :Calc_MetrologyBTS_Delta_MAX_MIN_AVG(Sensor_Num_StrA+Sensor_Num_StrB+..+Sensor_Num_StrN)
    Dim DSPWF_MTR_BTS_C0_Data As New DSPWave
    Dim DSPWF_DictData As New DSPWave
    Dim DSPWF_MTR_BTS_C0_MAX As New DSPWave
    Dim DSPWF_MTR_BTS_C0_MIN As New DSPWave
    Dim DSPWF_MTR_BTS_C0_AVG As New DSPWave
    Dim DSPWF_MTR_BTS_C0_Delta_MAX_AVG As New DSPWave
    Dim DSPWF_MTR_BTS_C0_Delta_MIN_AVG As New DSPWave
    Dim TestNameInput As String
    Dim l_SensorCnt As Long
    Dim vsite As Variant
    Dim SensorAry() As String: SensorAry = Split(argv(0), "+")
    
    If UBound(SensorAry) + 1 = 0 Then
        Call Print_Error_Message(Warning_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Delta_MAX_MIN_AVG", "No Sensor data be calculate!!")
        Exit Function
    Else
        DSPWF_MTR_BTS_C0_Data.CreateConstant 0, UBound(SensorAry) + 1, DspDouble
        DSPWF_MTR_BTS_C0_MAX.CreateConstant 0, 1, DspDouble
        DSPWF_MTR_BTS_C0_MIN.CreateConstant 0, 1, DspDouble
        DSPWF_MTR_BTS_C0_AVG.CreateConstant 0, 1, DspDouble
        DSPWF_MTR_BTS_C0_Delta_MAX_AVG.CreateConstant 0, 1, DspDouble
        DSPWF_MTR_BTS_C0_Delta_MIN_AVG.CreateConstant 0, 1, DspDouble
    End If
    
    For l_SensorCnt = 0 To UBound(SensorAry)
       ' DSPWF_DictData = GetStoredCaptureData(SensorAry(l_SensorCnt) & "_C0_SRC_DEC")
        DSPWF_DictData = GetStoredCaptureData(SensorAry(l_SensorCnt) & "_c0_cal_avg")
        For Each vsite In TheExec.sites
            DSPWF_MTR_BTS_C0_Data.Element(l_SensorCnt) = DSPWF_DictData.Element(0)
        Next vsite
    Next l_SensorCnt
    
    For Each vsite In TheExec.sites
        DSPWF_MTR_BTS_C0_MAX.Element(0) = FormatNumber(DSPWF_MTR_BTS_C0_Data(vsite).CalcMaximumValue, 4)
        DSPWF_MTR_BTS_C0_MIN.Element(0) = FormatNumber(DSPWF_MTR_BTS_C0_Data(vsite).CalcMinimumValue, 4)
        DSPWF_MTR_BTS_C0_AVG.Element(0) = FormatNumber(DSPWF_MTR_BTS_C0_Data(vsite).CalcMean, 4)
        DSPWF_MTR_BTS_C0_Delta_MAX_AVG.Element(0) = FormatNumber(DSPWF_MTR_BTS_C0_MAX.Element(0) - DSPWF_MTR_BTS_C0_AVG.Element(0), 4)
        DSPWF_MTR_BTS_C0_Delta_MIN_AVG.Element(0) = FormatNumber(DSPWF_MTR_BTS_C0_MIN.Element(0) - DSPWF_MTR_BTS_C0_AVG.Element(0), 4)
    Next vsite
    '--- Print datalog ---
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, "C0-Maximum", , , , , , tlForceNone)
    TheExec.flow.TestLimit resultVal:=DSPWF_MTR_BTS_C0_MAX.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, lowVal:=-9999, hiVal:=9999
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, "C0-Minimum", , , , , , tlForceNone)
    TheExec.flow.TestLimit resultVal:=DSPWF_MTR_BTS_C0_MIN.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, lowVal:=-9999, hiVal:=9999
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, "C0-Average", , , , , , tlForceNone)
    TheExec.flow.TestLimit resultVal:=DSPWF_MTR_BTS_C0_AVG.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, lowVal:=-9999, hiVal:=9999
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, "C0-MAX-AVG", , , , , , tlForceNone)
    TheExec.flow.TestLimit resultVal:=DSPWF_MTR_BTS_C0_Delta_MAX_AVG.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, lowVal:=0, hiVal:=10
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, "C0-MIN-AVG", , , , , , tlForceNone)
    TheExec.flow.TestLimit resultVal:=DSPWF_MTR_BTS_C0_Delta_MIN_AVG.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, lowVal:=-10, hiVal:=0
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Delta_MAX_MIN_AVG")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_ICTS_Coefficient(argc As Integer, argv() As String) As Long
    '20240131 Hidra by YM
    
    On Error GoTo errHandler
    
    Dim TestNameInput As String
    Dim x0 As New SiteDouble
    Dim Xnb As New SiteDouble
    Dim a3, a2, a1, a0, k0, k1, m0, m1, m2, Tcp As New SiteDouble
    Dim Sensor_Num_Str As String: Sensor_Num_Str = argv(2)
    
    x0 = GetStoreDataAllType(argv(0) & "_para")
    x0 = x0.divide(2 ^ 16)
    
    Xnb = GetStoreDataAllType(argv(1) & "_para")
    Xnb = Xnb.divide(2 ^ 16)
    
    If TheExec.TesterMode = testModeOffline Then
        x0 = 0.0005198
    End If
 
    
    'ICTS Golden coefficients (will be in Table):
    'a3 = 0
    'a2 = -16.30924774
    'a1 = 324.13354724
    'a0 = 49.63661904
    'k0 = 324.98372688
    'k1 = 16.47506753
    'm0 = 324.78353753
    'm1 = 17.4639996
    'm2 = 22.70347861
    'Tcp = 25#
    
    a3 = 0
    a2 = -16.30924774
    a1 = 324.13354724
    a0 = 49.63661904
    k0 = 324.98372688
    k1 = 16.47506753
    m0 = 324.78353753
    m1 = 17.4639996
    m2 = 22.70347861
    Tcp = 25
    
    TheExec.Datalog.WriteComment "********************** Coefficients for Calibration **********************"
    TheExec.Datalog.WriteComment " a3  = " & a3
    TheExec.Datalog.WriteComment " a2  = " & a2
    TheExec.Datalog.WriteComment " a1  = " & a1
    TheExec.Datalog.WriteComment " a0  = " & a0
    TheExec.Datalog.WriteComment " k0  = " & k0
    TheExec.Datalog.WriteComment " k1  = " & k1
    TheExec.Datalog.WriteComment " m0  = " & m0
    TheExec.Datalog.WriteComment " m1  = " & m1
    TheExec.Datalog.WriteComment " m2  = " & m2
    TheExec.Datalog.WriteComment " Tcp = " & Tcp
    TheExec.Datalog.WriteComment "**************************************************************************"
    
    'Calculate Coefficients for eFuse
    'C0 = Tcp - a1 * x0 - a2 * x0^2
    'C1 = a1
    'C2 = a2
    'C3 = 0
    'C1a = (k0 + k1 * Xnb ), Xnb from T2.1 Xnb = 0, If T2.1 not done.
    'C0a = Tcp - c1a * x0 - a2 * x0^2
    
    Dim C0 As New SiteDouble
    Dim C1 As New SiteDouble
    Dim C2 As New SiteDouble
    Dim C3 As New SiteDouble
    Dim C1a As New SiteDouble
    Dim C0a As New SiteDouble
    
    C0 = x0.Multiply(a1).Negate.Add(Tcp).Subtract(x0.power(2).Multiply(a2))
    C1 = a1
    C2 = a2
    C3 = 0
    C1a = Xnb.Multiply(k1).Add(k0)
    C0a = x0.Multiply(C1a).Negate.Add(Tcp).Subtract(x0.power(2).Multiply(a2))
    
    Dim DSP_C0_Cal_eFuse As New DSPWave
    Dim DSP_C1_Cal_eFuse As New DSPWave
    Dim DSP_C2_Cal_eFuse As New DSPWave
    Dim DSP_C3_Cal_eFuse As New DSPWave
    Dim DSP_C0a_Cal_eFuse As New DSPWave
    Dim DSP_C1a_Cal_eFuse As New DSPWave
    Dim DSP_C0_CAL  As New DSPWave
    Dim DSP_C1_CAL  As New DSPWave
    Dim DSP_C2_CAL  As New DSPWave
    Dim DSP_C3_CAL  As New DSPWave
    Dim DSP_C0a_CAL As New DSPWave
    Dim DSP_C1a_CAL As New DSPWave
    Dim DSP_C0_Cal_Src As New DSPWave
    Dim DSP_C1_Cal_Src As New DSPWave
    Dim DSP_C2_Cal_Src As New DSPWave
    Dim DSP_C3_Cal_Src As New DSPWave
    Dim DSP_C0a_Cal_Src As New DSPWave
    Dim DSP_C1a_Cal_Src As New DSPWave
        
    Set DSP_C0_Cal_eFuse = Nothing
    Set DSP_C1_Cal_eFuse = Nothing
    Set DSP_C2_Cal_eFuse = Nothing
    Set DSP_C3_Cal_eFuse = Nothing
    Set DSP_C0a_Cal_eFuse = Nothing
    Set DSP_C1a_Cal_eFuse = Nothing
    Set DSP_C0_CAL = Nothing
    Set DSP_C1_CAL = Nothing
    Set DSP_C2_CAL = Nothing
    Set DSP_C3_CAL = Nothing
    Set DSP_C0a_CAL = Nothing
    Set DSP_C1a_CAL = Nothing
        
    DSP_C0_Cal_eFuse.CreateConstant 0, 1
    DSP_C1_Cal_eFuse.CreateConstant 0, 1
    DSP_C2_Cal_eFuse.CreateConstant 0, 1
    DSP_C3_Cal_eFuse.CreateConstant 0, 1
    DSP_C0a_Cal_eFuse.CreateConstant 0, 1
    DSP_C1a_Cal_eFuse.CreateConstant 0, 1
    DSP_C0_CAL.CreateConstant 0, 1
    DSP_C1_CAL.CreateConstant 0, 1
    DSP_C2_CAL.CreateConstant 0, 1
    DSP_C3_CAL.CreateConstant 0, 1
    DSP_C0a_CAL.CreateConstant 0, 1
    DSP_C1a_CAL.CreateConstant 0, 1
    
    For Each site In TheExec.sites.Active
        DSP_C0_CAL.Element(0) = C0
        DSP_C1_CAL.Element(0) = C1
        DSP_C2_CAL.Element(0) = C2
        DSP_C3_CAL.Element(0) = C3
        DSP_C0a_CAL.Element(0) = C0a
        DSP_C1a_CAL.Element(0) = C1a
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C0_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling

    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C1_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling

    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C2_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling

    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C3_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling

    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C0a_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling

    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_C1a_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling

    C0 = C0.Multiply(2 ^ 9)
    C1 = C1.Multiply(2 ^ 9)
    C2 = C2.Multiply(2 ^ 9)
    C3 = C3.Multiply(2 ^ 9)
    C0a = C0a.Multiply(2 ^ 9)
    C1a = C1a.Multiply(2 ^ 9)
    For Each site In TheExec.sites.Active
        DSP_C0_CAL.Element(0) = C0
        DSP_C1_CAL.Element(0) = C1
        DSP_C2_CAL.Element(0) = C2
        DSP_C3_CAL.Element(0) = C3
        DSP_C0a_CAL.Element(0) = C0a
        DSP_C1a_CAL.Element(0) = C1a
    Next site
        
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0_Cal_eFuse, DSP_C0_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1_Cal_eFuse, DSP_C1_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C2_Cal_eFuse, DSP_C2_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C3_Cal_eFuse, DSP_C3_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0a_Cal_eFuse, DSP_C0a_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1a_Cal_eFuse, DSP_C1a_CAL, 20, 0)
        
    Set DSP_C0_Cal_Src = Nothing
    Set DSP_C1_Cal_Src = Nothing
    Set DSP_C2_Cal_Src = Nothing
    Set DSP_C3_Cal_Src = Nothing
    Set DSP_C0a_Cal_Src = Nothing
    Set DSP_C1a_Cal_Src = Nothing
        
    Call HardIP_Dec2Bin(DSP_C0_Cal_Src, DSP_C0_Cal_eFuse, 20) 'DSP_DecToBin
    Call HardIP_Dec2Bin(DSP_C1_Cal_Src, DSP_C1_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C2_Cal_Src, DSP_C2_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C3_Cal_Src, DSP_C3_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C0a_Cal_Src, DSP_C0a_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C1a_Cal_Src, DSP_C1a_Cal_eFuse, 20)
        
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites.Active
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C0=" & DSP_C0_CAL.Element(0) & " => " & DSP_C0_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C1=" & DSP_C1_CAL.Element(0) & " => " & DSP_C1_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C2=" & DSP_C2_CAL.Element(0) & " => " & DSP_C2_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C3=" & DSP_C3_CAL.Element(0) & " => " & DSP_C3_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C0A=" & DSP_C0a_CAL.Element(0) & " => " & DSP_C0a_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num_Str & "_C1A=" & DSP_C1a_CAL.Element(0) & " => " & DSP_C1a_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment ""
        Next site
    End If
    
    Call StoreDataAllType(Sensor_Num_Str & "_C0_SRC", DSP_C0_Cal_Src)
        
    Call StoreDataAllType(Sensor_Num_Str & "_C1_SRC", DSP_C1_Cal_Src)
        
    Call StoreDataAllType(Sensor_Num_Str & "_C2_SRC", DSP_C2_Cal_Src)
        
    Call StoreDataAllType(Sensor_Num_Str & "_C3_SRC", DSP_C3_Cal_Src)
    
    Call StoreDataAllType(Sensor_Num_Str & "_C0A_SRC", DSP_C0a_Cal_Src)
    
    Call StoreDataAllType(Sensor_Num_Str & "_C1A_SRC", DSP_C1a_Cal_Src)
      
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_ICTS_Coefficient") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_ICTS_T2P1A(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim DSP_Temperature As New DSPWave
    Dim Temperature_AVG As New SiteDouble
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    
    'a.  REF_TEMP where REF_TEMP =  (Temp_BTS_Te006+Temp_BTS_Te007)/2
    DSP_Temperature.CreateConstant 0, UBound(Temperature_Sensor) + 1, DspDouble
    For Each site In TheExec.sites
        For i = 0 To UBound(Temperature_Sensor)
            Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
            Temperature = Temperature.divide(64)
            DSP_Temperature.Element(i) = Temperature
        Next i
        Temperature_AVG = DSP_Temperature.CalcMean
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=Temperature_AVG, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'b.  ICTS_NB_TC1 = (X_NB_NOW - X_NB_CP1) / (REF_TEMP - 25.0)
    Dim ICTS_NB_TC1 As New SiteDouble
    Dim X_NB_CP1 As New SiteDouble
    Dim DSP_X_NB_CP1 As New DSPWave
    Dim DSP_X_NB_CP1_DEC As New DSPWave
    Dim Bit_Size As New SiteLong
    Dim X_NB_NOW As New SiteDouble
    
    DSP_X_NB_CP1_DEC.CreateConstant 0, 1, DspLong
    DSP_X_NB_CP1 = GetStoreDataAllType(argv(1))
    
    For Each site In TheExec.sites
        Bit_Size = DSP_X_NB_CP1.SampleSize
    Next site
    
    rundsp.DSP_2S_Complement_To_SignDec DSP_X_NB_CP1, Bit_Size, DSP_X_NB_CP1_DEC
    
    
    
    
    
    For Each site In TheExec.sites
        X_NB_CP1 = DSP_X_NB_CP1_DEC.Element(0)
    Next site
    
    X_NB_NOW = GetStoreDataAllType(argv(2) + "_para")
    X_NB_CP1 = X_NB_CP1.divide(2 ^ 16)
    X_NB_NOW = X_NB_NOW.divide(2 ^ 16)
    
    '2.  Log X_NB_CP1
    '4.  Log X_NB_NOW
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites.Active
            TheExec.Datalog.WriteComment "site_" & site & "_" & "X_NB_CP1 = " & X_NB_CP1
            TheExec.Datalog.WriteComment "site_" & site & "_" & "X_NB_NOW = " & X_NB_NOW
        Next site
    End If
    
    
    For Each site In TheExec.sites
        If Temperature_AVG(site) = 25 Then
            Temperature_AVG(site) = 25.0000001
        End If
    Next site
    
    ICTS_NB_TC1 = X_NB_NOW.Subtract(X_NB_CP1).divide(Temperature_AVG.Subtract(25))
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=ICTS_NB_TC1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'c.  ICTS_TC1_NB_RATIO = ICTS_NB_TC1 / (1 + 0.5 * X_NB_CP1)
    Dim ICTS_TC1_NB_RATIO As New SiteDouble
    ICTS_TC1_NB_RATIO = ICTS_NB_TC1.divide(X_NB_CP1.Multiply(0.5).Add(1))
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=ICTS_TC1_NB_RATIO, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'd.  ICTS_NB_SHIFT = X_NB_NOW - X_NB_CP1
    Dim ICTS_NB_SHIFT As New SiteDouble
    ICTS_NB_SHIFT = X_NB_NOW.Subtract(X_NB_CP1)
       
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=ICTS_NB_SHIFT, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_ICTS_T2P1A") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_ICTS_T2P2A(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'Alg::Calc_ICTS_T2P2A(bts__ts000_m1_temp10+bts__ts001_m1_temp10,icts_res_tst_ric_ate,icts_res_tst_rhim_ate,icts_res_tst_RIC_NOW_ate_para,icts_res_tst_RH_NOW_ate_para)
    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim DSP_Temperature As New DSPWave
    Dim Temperature_AVG As New SiteDouble
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    
    '2) REF_TEMP where REF_TEMP =  (Temp_BTS_Ts000+Temp_BTS_Ts001)/2
    DSP_Temperature.CreateConstant 0, UBound(Temperature_Sensor) + 1, DspDouble
    For Each site In TheExec.sites
        For i = 0 To UBound(Temperature_Sensor)
            Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
            Temperature = Temperature.divide(64)
            DSP_Temperature.Element(i) = Temperature
        Next i
        Temperature_AVG = DSP_Temperature.CalcMean
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=Temperature_AVG, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    
    Dim RIC_TC1 As New SiteDouble
    Dim RIC_CP1 As New SiteDouble
    Dim RIC_NOW As New SiteDouble
    Dim DSP_RIC_CP1 As New DSPWave
    Dim DSP_RIC_CP1_DEC As New DSPWave
    
    Dim RH_TC1 As New SiteDouble
    Dim RH_CP1 As New SiteDouble
    Dim RH_NOW As New SiteDouble
    Dim DSP_RH_CP1 As New DSPWave
    Dim DSP_RH_CP1_DEC As New DSPWave
    
    DSP_RIC_CP1 = GetStoreDataAllType(argv(1))  'icts_res_tst_ric_ate
    DSP_RH_CP1 = GetStoreDataAllType(argv(2))   'icts_res_tst_rhim_ate
    
    DSP_RIC_CP1_DEC.CreateConstant 0, 1, DspLong
    DSP_RH_CP1_DEC.CreateConstant 0, 1, DspLong
 
    rundsp.BinToDec DSP_RIC_CP1, DSP_RIC_CP1_DEC
    rundsp.BinToDec DSP_RH_CP1, DSP_RH_CP1_DEC
    
    'Put DSP into Sitedouble
    For Each site In TheExec.sites
        RIC_CP1 = DSP_RIC_CP1_DEC.Element(0)
        RH_CP1 = DSP_RH_CP1_DEC.Element(0)
    Next site
    
    RIC_NOW = GetStoreDataAllType(argv(3))     'icts_res_tst_RIC_NOW_ate_para
    RH_NOW = GetStoreDataAllType(argv(4))     'icts_res_tst_RH_NOW_ate_para
    
    For Each site In TheExec.sites
        RIC_CP1 = DSP_RIC_CP1_DEC.Element(0)
        RH_CP1 = DSP_RH_CP1_DEC.Element(0)
        '20240515 Hidra by YM -- Prevent Read Zero fuse
        If RIC_CP1 = 0 Then
            RIC_CP1 = 0.00000001
            TheExec.Datalog.WriteComment "error in Calc_ICTS_T2P2A RIC_CP1 value is zero!!! "
        End If
        If RH_CP1 = 0 Then
            RH_CP1 = 0.00000001
            TheExec.Datalog.WriteComment "error in Calc_ICTS_T2P2A RH_CP1 value is zero!!! "
        End If
        If RIC_NOW = 0 Then
            RIC_NOW = 0.00000001
            TheExec.Datalog.WriteComment "error in Calc_ICTS_T2P2A RIC_NOW value is zero!!! "
        End If
        If RH_NOW = 0 Then
            RH_NOW = 0.00000001
            TheExec.Datalog.WriteComment "error in Calc_ICTS_T2P2A RH_NOW value is zero!!! "
        End If
        '20240515 Hidra by YM -- Prevent Read Zero fuse
    Next site
    
    'a.RIC_TC1 = (RIC_NOW - RIC_CP1) / (REF_TEMP - 25#)
    'd.RH_TC1 = (RH_NOW - RH_CP1) / (REF_TEMP - 25#)
    
    
    For Each site In TheExec.sites
        If Temperature_AVG(site) = 25 Then
            Temperature_AVG(site) = 25.0000001
        End If
    Next site
    
    
    RIC_TC1 = RIC_NOW.Subtract(RIC_CP1).divide(Temperature_AVG.Subtract(25))
    RH_TC1 = RH_NOW.Subtract(RH_CP1).divide(Temperature_AVG.Subtract(25))
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RIC_TC1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RH_TC1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'b.RIC_TC1_R_RATIO = RIC_TC1 / RIC_CP1
    'e.RH_TC1_R_RATIO = RH_TC1 / RH_CP1
    Dim RIC_TC1_R_RATIO As New SiteDouble
    Dim RH_TC1_R_RATIO As New SiteDouble
    
    RIC_TC1_R_RATIO = RIC_TC1.divide(RIC_CP1)
    RH_TC1_R_RATIO = RH_TC1.divide(RH_CP1)
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RIC_TC1_R_RATIO, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RH_TC1_R_RATIO, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'c.RIC_SHIFT = RIC_NOW - RIC_CP1
    'f.RH_SHIFT = RH_NOW - RH_CP1
    Dim RIC_SHIFT As New SiteDouble
    Dim RH_SHIFT As New SiteDouble
    
    RIC_SHIFT = RIC_NOW.Subtract(RIC_CP1)
    RH_SHIFT = RH_NOW.Subtract(RH_CP1)
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RIC_SHIFT, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RH_SHIFT, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'g.RR_NOW = RIC_NOW / RH_NOW
    'h.RR_CP1 = RIC_CP1 / RH_CP1
    Dim RR_NOW As New SiteDouble
    Dim RR_CP1 As New SiteDouble
    
    RR_NOW = RIC_NOW.divide(RH_NOW)
    RR_CP1 = RIC_CP1.divide(RH_CP1)
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RR_NOW, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RR_CP1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'i.RR_TC1 = (RR_NOW - RR_CP1) / (REF_TEMP - 25#)
    Dim RR_TC1 As New SiteDouble
    
    RR_TC1 = RR_NOW.Subtract(RR_CP1).divide(Temperature_AVG.Subtract(25))
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RR_TC1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'j.RR_TC1_R_RATIO = RR_TC1 / RR_CP1
    Dim RR_TC1_R_RATIO As New SiteDouble
    
    RR_TC1_R_RATIO = RR_TC1.divide(RR_CP1)
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RR_TC1_R_RATIO, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    'k.RR_SHIFT = RR_NOW - RR_CP1
    Dim RR_SHIFT As New SiteDouble
    
    RR_SHIFT = RR_NOW.Subtract(RR_CP1)
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=RR_SHIFT, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_ICTS_T2P2A") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MetrologyBTS_Delta(argc As Integer, argv() As String) As Long
On Error GoTo errHandler

    Dim sd_DictData0 As New SiteDouble
    Dim sd_DictData1 As New SiteDouble
    Dim sd_Temp_Delta As New SiteDouble
    Dim TestNameInput As String
    Dim l_SensorCnt As Long
    Dim SensorAry() As String: SensorAry = Split(argv(0), "+")
    
    If UBound(SensorAry) + 1 = 0 Then
        Call Print_Error_Message(Warning_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Delta", "No Sensor data be calculate!!")
        Exit Function
    End If
    
    For l_SensorCnt = 0 To UBound(SensorAry)
        sd_DictData0 = GetStoredData(SensorAry(l_SensorCnt) & "_" & argv(1))
        sd_DictData1 = GetStoredData(SensorAry(l_SensorCnt) & "_" & argv(2))
        sd_Temp_Delta = sd_DictData0.Subtract(sd_DictData1)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, argv(1) & "-" & argv(2), , , , , , tlForceFlow)
        TheExec.flow.TestLimit resultVal:=sd_Temp_Delta, Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next l_SensorCnt

  
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Delta")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_MetrologyBTS_Temperature_Delta(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim Temperature As New SiteDouble
    Dim Temperature_T4P1C As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim Temperature_Array_Delta(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_Temperature_Delta() As New DSPWave: ReDim DSP_Temperature_Delta(UBound(Temperature_Sensor))
    
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    Dim Sensor_Num() As String
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
        Temperature_T4P1C = GetStoreDataAllType(Replace(LCase(Temperature_Sensor(i)), "_mn", "") + "_para")
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).data = Temperature_Array
            
            Temperature_Array_Delta(0) = (Temperature - Temperature_T4P1C) / 64
            DSP_Temperature_Delta(i).data = Temperature_Array_Delta
            
        Next site
    Next i
    
    For i = 0 To UBound(Temperature_Sensor)
        Sensor_Num = Split(Temperature_Sensor(i), "_")
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
    For i = 0 To UBound(Temperature_Sensor)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=DSP_Temperature_Delta(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Temperature") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_NTS_Statistics(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    'argv(0):tg00i__vref_para@tg00i__vptatp_para+tg00j__vref_para@tg00j__vptatp_para    =>To calc delta '@' means delta ; '+' to join multi data
    'argv(1):VREF            => Calc avg,max,min,std,max-avg,min-avg
    'argv(2):VPTATP          => Calc avg,max,min,std,max-avg,min-avg
    'argv(3):12B_1X          => Calc avg,max,min,std,max-avg,min-avg
    'argv(4):12B_1X_SELF_0P5 => Calc avg,max,min,std,max-avg,min-avg
    
    'Delta Var
    Dim VREF As New SiteDouble
    Dim VPTATP As New SiteDouble
    Dim delta_Ary() As String: delta_Ary = Split(argv(0), "+")
    Dim sensor_Ary() As String
    Dim PLD_Data_Delta As New PinListData
    Dim Avg_Delta As New SiteDouble
    Dim Max_Delta As New SiteDouble
    Dim Min_Delta As New SiteDouble
    Dim Std_Delta As New SiteDouble
    'Statistics Var
    Dim PLD_Ary() As String
    Dim sensor As String
    Dim PLD_Data() As New PinListData: ReDim PLD_Data(argc - 1)
    Dim Avg() As New SiteDouble: ReDim Avg(argc - 1)
    Dim Max() As New SiteDouble: ReDim Max(argc - 1)
    Dim Min() As New SiteDouble: ReDim Min(argc - 1)
    Dim Std() As New SiteDouble: ReDim Std(argc - 1)
    Dim i, j As Long
    Dim site As Variant
    Dim TestNameInput As String
    
    'Calc (VREF-VPTATP) Per Sensor
    For i = 0 To UBound(delta_Ary)
        sensor_Ary = Split(delta_Ary(i), "@")   'tg00i__vref_para@tg00i__vptatp_para
        VREF = GetStoreDataAllType(sensor_Ary(0))   'tg00i__vref_para
        VPTATP = GetStoreDataAllType(sensor_Ary(1)) 'tg00i__vptatp_para
        PLD_Data_Delta.AddPin (delta_Ary(i))
        PLD_Data_Delta.pins(delta_Ary(i)) = VREF.Subtract(VPTATP)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=PLD_Data_Delta.pins(delta_Ary(i)), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    'Calc Statistics for VREF-VPTATP
    Avg_Delta = PLD_Data_Delta.Analyze.mean
    Max_Delta = PLD_Data_Delta.Analyze.Maximum
    Min_Delta = PLD_Data_Delta.Analyze.Minimum
    Std_Delta = PLD_Data_Delta.Analyze.stdDev
    'Avg
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=Avg_Delta, Tname:=TestNameInput, ForceResults:=tlForceFlow
    'Min
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=Min_Delta, Tname:=TestNameInput, ForceResults:=tlForceFlow
    'Max
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=Max_Delta, Tname:=TestNameInput, ForceResults:=tlForceFlow
    'Max-Avg
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=Max_Delta.Subtract(Avg_Delta), Tname:=TestNameInput, ForceResults:=tlForceFlow
    'Min-Avg
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=Min_Delta.Subtract(Avg_Delta), Tname:=TestNameInput, ForceResults:=tlForceFlow
    'STD
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=Std_Delta, Tname:=TestNameInput, ForceResults:=tlForceFlow
    'Calc Statistics for VREF-VPTATP
    
    
    'Calc Statistics
    For i = 0 To argc - 2
        PLD_Ary = Split(argv(i + 1), "+")
        For j = 0 To UBound(PLD_Ary)
            sensor = PLD_Ary(j)
            PLD_Data(i).AddPin (sensor)
            PLD_Data(i).pins(sensor) = GetStoreDataAllType(sensor)
        Next j
        Avg(i) = PLD_Data(i).Analyze.mean
        Max(i) = PLD_Data(i).Analyze.Maximum
        Min(i) = PLD_Data(i).Analyze.Minimum
        Std(i) = PLD_Data(i).Analyze.stdDev
        
        'Avg
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Avg(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        'Min
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Min(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        'Max
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Max(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        'Max-Avg
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Max(i).Subtract(Avg(i)), Tname:=TestNameInput, ForceResults:=tlForceFlow
        'Min-Avg
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Min(i).Subtract(Avg(i)), Tname:=TestNameInput, ForceResults:=tlForceFlow
        'STD
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.flow.TestLimit resultVal:=Std(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_NTS_Statistics") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

