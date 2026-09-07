Attribute VB_Name = "LIB_HardIP_Metrology_AP"
Option Explicit 'Add ErrHandler 2023/05/29
'Type MTRSNS_Matrix
'    ROT_Matrix() As Double
'    ROV_Matrix() As Double
'    ROT_a_max_min_Matrix() As Double
'    ROV_a_max_min_Matrix() As Double
'End Type
'Public MetrologySense_Matrix() As MTRSNS_Matrix
Public Flag_TMPS_1st_Run As Boolean

Public Function MetrologyTMPS_Measurement_Process(Pat As String, srcPin As PinList, code() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef Rtn_MeasVolt() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, digsrc_assignment As String, TrimStoreName() As String, MeasV_WaitTime As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim srcWave() As New DSPWave: ReDim srcWave(UBound(code))
    Dim site As Variant
    Dim InDSPWave As New DSPWave
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
    
    Call GeneralDigSrcSettingWithBurst(Pat, srcPin, InDSPWave)

    TheHdw.Patterns(Pat).start
    
    For i = 0 To NumberOfMeasV - 1
        TheHdw.Digital.Patgen.FlagWait cpuA, 0
        Rtn_MeasVolt(i) = HardIP_MeasureVolt
        Call DebugPrintFunc_PPMU(vbNullString)
        
        'Oscar, DiffMeter
                
        For Each site In TheExec.sites
            Res(i) = Abs(Rtn_MeasVolt(i).Pins(0).value - Rtn_MeasVolt(i).Pins(1).value)
                        'Res(i) = Abs(Rtn_MeasVolt(i).Pins(0).value)
            For j = 0 To Rtn_MeasVolt(i).Pins.Count - 1
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i) & ", Pin : " & Rtn_MeasVolt(i).Pins(j) & ", Voltage = " & Rtn_MeasVolt(i).Pins(j).value
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
    Dim site As Variant 'Carter, 20240304
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
    TheExec.Flow.TestLimit resultVal:=DSP_OffSet_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
    Call StoreDataAllType(argv(1) & "_" & CStr(TheExec.Flow.var(argv(2)).value), DSP_OffSet_Mean)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_OffSet_Mean_eFuse, DSP_OffSet_Mean, 18, 0)
    Call StoreDataAllType(argv(1) & "_eFuse_" & CStr(TheExec.Flow.var(argv(2)).value), DSP_OffSet_Mean_eFuse)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyTMPS_OffSet") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
Dim site As Variant 'Carter, 20240304
For i = 0 To UBound(InWf_Array)
    InWf_SiteDouble(i) = GetStoreDataAllType(InWf_Split(i) & "_para")
Next i

DSP_OffSet_Mean = GetStoreDataAllType(argv(1) & "_" & CStr(TheExec.Flow.var(argv(2)).value))
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
TheExec.Flow.TestLimit resultVal:=DSP_Gain_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

Call StoreDataAllType(argv(3) & "_" & CStr(TheExec.Flow.var(argv(2)).value), DSP_Gain_Mean)

If TheExec.Flow.var(argv(2)).value = argv(5) Then
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
    TheExec.Flow.TestLimit resultVal:=DSP_Gain_Mean_Final.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
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
Dim site As Variant 'Carter, 20240304
For Each site In TheExec.sites.Active
    DSP_DSSCOUT_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0)
    DSP_DSSCOUT_Mean.data = DSP_DSSCOUT_Mean_Array
Next site

TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
TheExec.Flow.TestLimit resultVal:=DSP_DSSCOUT_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
Call StoreDataAllType(StoreName_DSSC_OUT_Mean, DSP_DSSCOUT_Mean)

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "MTRTMPS_DSSCOUT_AVG") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MetrologyGR_Measurement_Process(Pat As String, srcPin As PinList, code As SiteLong, Res As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef Rtn_MeasVolt() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, digsrc_assignment As String, TrimStoreName As String, MeasV_WaitTime As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim sigName As String, srcWave As New DSPWave, site As Variant
    Dim InDSPWave As New DSPWave
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
    
    Call GeneralDigSrcSettingWithBurst(LCase(Pat), srcPin, InDSPWave)
   
    TheHdw.Patterns(Pat).start
    
    For i = 0 To NumberOfMeasV - 1
        Instance_Data.TestSeqNum = i
        TheHdw.Digital.Patgen.FlagWait cpuA, 0

        Rtn_MeasVolt(i) = HardIP_MeasureVolt
        Call DebugPrintFunc_PPMU(vbNullString)

        For Each site In TheExec.sites

            If i = NumberOfMeasV - 1 Then Res = Abs(Rtn_MeasVolt(i).Pins(0).value - Rtn_MeasVolt(i - 1).Pins(0).value)
            For j = 0 To Rtn_MeasVolt(i).Pins.Count - 1
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Pin : " & Rtn_MeasVolt(i).Pins(j) & ", Voltage = " & Rtn_MeasVolt(i).Pins(j).value
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
        DSP_T1 = DSP_APM_Count_Out_0.ConvertStreamTo(tldspParallel, DSP_APM_Count_Out_0.sampleSize, 0, Bit0IsMsb).Multiply(375000).Add(0.001).Reciprocate
        DSP_T2 = DSP_APM_Count_Out_1.ConvertStreamTo(tldspParallel, DSP_APM_Count_Out_1.sampleSize, 0, Bit0IsMsb).Multiply(375000).Add(0.001).Reciprocate
        DSP_TRC = DSP_T2.Subtract(DSP_T1)
        DSP_TRC_ERR = DSP_TRC.Subtract(0.0000000017).divide(0.0000000017).Multiply(100).divide(0.5)
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_T1.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_T2.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_TRC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_TRC_ERR.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
        Call StoreDataAllType(SubBlockName & "-" & Instance_Data.Tname(TheExec.Flow.TestLimitIndex), MetrologySense_Frequency(i))
        TestNameInput = Report_TName_From_Instance("CalcF", vbNullString)
        TheExec.Flow.TestLimit resultVal:=MetrologySense_Frequency(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).data = Temperature_Array
        Next site
    Next i
    If argc = 2 Then
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
    End If
    
    For i = 0 To UBound(Temperature_Sensor)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.Flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    If TheExec.Flow.enableWord("TMPS_Monitor") = True Then
        TheHdw.Pins("all_digital").Digital.InitState = chInitLo
        TheHdw.DCVS.Pins("VDD_SRAM_GPU,VDD_GPU,VDD_ECPU,VDD_PCPU,VDD_CPU_SRAM").Voltage.Main.value = 0.5
        If Not (Flag_TMPS_1st_Run) Then
            TheHdw.Wait 5
            Flag_TMPS_1st_Run = True
        Else
            TheHdw.Wait 1
        End If
    End If
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
    
    TheExec.Flow.TestLimit resultVal:=SiteDbl_Vref_Error, lowVal:=Low_limit, hiVal:=High_limit, Tname:=TestNameInput, ForceResults:=tlForceFlow

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
    Dim A0 As Double
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
    Dim site As Variant 'Carter, 20240304
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
        A0 = -179.0341695
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
        A0 = -208.7079186
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
        A0 = -208.7079186
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
        A0 = -208.7079186
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
        A0 = -208.7079186
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
        A0 = -59.285102
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
        A0 = -59.285102
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
    TheExec.Flow.TestLimit resultVal:=DSP_ADC_Temperature_Sensor_Raw_Data.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_coeff0", "_", vbNullString))
    TheExec.Flow.TestLimit resultVal:=DSP_Coefficient_C0.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_coeff1", "_", vbNullString))
    TheExec.Flow.TestLimit resultVal:=DSP_Coefficient_C1.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_coeff2", "_", vbNullString))
    TheExec.Flow.TestLimit resultVal:=DSP_Coefficient_C2.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, Replace(argv(3) & "_coeff3", "_", vbNullString))
    TheExec.Flow.TestLimit resultVal:=DSP_Coefficient_C3.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow

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
Dim Sensor As String: Sensor = argv(1)
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
        MetrologySense_ROT_Frequency = GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT")
        MetrologySense_ROV_Frequency = GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV")
    Else
        For Each site In TheExec.sites
            MetrologySense_ROT_Frequency = MetrologySense_ROT_Frequency.Concatenate(GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT"))
            MetrologySense_ROV_Frequency = MetrologySense_ROV_Frequency.Concatenate(GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV"))
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
    
    
    
    DSPTempCal_1 = DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row).Copy
    DSPTempCal_2 = DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).Copy
    DSPTempCal_3 = DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row).Copy
    a1_Compression = a1_LogicalCompare.Subtract(DSPTempCal_1).divide(DSPTempCal_2.Subtract(DSPTempCal_3))
    
    DSPTempCal_1 = DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row).Copy
    DSPTempCal_2 = DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).Copy
    DSPTempCal_3 = DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row).Copy
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
    TheExec.Flow.TestLimit resultVal:=a1.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
    TheExec.Flow.TestLimit resultVal:=a1_Compression.Element(i), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Tname:=TestNameInput, ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        a1_Compression_eFuse_Store(i) = a1_Compression_eFuse.Select(i, , 1).Copy
    Next site
    Call StoreDataAllType("mtr_" & Sensor & "_t1_a1_" & CStr(i + 1), a1_Compression_eFuse_Store(i))
Next i
For i = 0 To MTRSNS_Matrix_ROV_Row - 1
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
    TheExec.Flow.TestLimit resultVal:=a2.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
    TheExec.Flow.TestLimit resultVal:=a2_Compression.Element(i), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Tname:=TestNameInput, ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        a2_Compression_eFuse_Store(i) = a2_Compression_eFuse.Select(i, , 1).Copy
    Next site
    Call StoreDataAllType("mtr_" & Sensor & "_t1_a2_" & CStr(i + 1), a2_Compression_eFuse_Store(i))
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
    Dim Sensor As String: Sensor = argv(1)
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
            MetrologySense_ROT_Frequency = GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT")
            MetrologySense_ROV_Frequency = GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV")
        Else
            For Each site In TheExec.sites
                MetrologySense_ROT_Frequency = MetrologySense_ROT_Frequency.Concatenate(GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT"))
                MetrologySense_ROV_Frequency = MetrologySense_ROV_Frequency.Concatenate(GetStoreDataAllType(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV"))
            Next site
        End If
    Next i
    For i = 0 To MTRSNS_Matrix_ROT_Row - 1
        If i = 0 Then
            a1_Compression_eFuse = GetStoreDataAllType("mtr_" & Sensor & "_t1_a1_" & CStr(i + 1))
        Else
            For Each site In TheExec.sites
                a1_Compression_eFuse = a1_Compression_eFuse.Concatenate(GetStoreDataAllType("mtr_" & Sensor & "_t1_a1_" & CStr(i + 1)))
            Next site
        End If
    Next i
    For i = 0 To MTRSNS_Matrix_ROV_Row - 1
        If i = 0 Then
            a2_Compression_eFuse = GetStoreDataAllType("mtr_" & Sensor & "_t1_a2_" & CStr(i + 1))
        Else
            For Each site In TheExec.sites
                a2_Compression_eFuse = a2_Compression_eFuse.Concatenate(GetStoreDataAllType("mtr_" & Sensor & "_t1_a2_" & CStr(i + 1)))
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
        a1_DeCompression = a1_Compression.Multiply(DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).Subtract(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row))).Add(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row).Copy)
        a2_DeCompression = a2_Compression.Multiply(DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).Subtract(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row))).Add(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row).Copy)
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
        TheExec.Flow.TestLimit resultVal:=MetrologySense_ROT_Frequency_DeCompression.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcF", vbNullString, , CInt(i))
        TheExec.Flow.TestLimit resultVal:=MetrologySense_ROV_Frequency_DeCompression.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
        TheExec.Flow.TestLimit resultVal:=MetrologySense_ROT_Frequency_Error.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , CInt(i))
        TheExec.Flow.TestLimit resultVal:=MetrologySense_ROV_Frequency_Error.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
    Dim site As Variant 'Carter, 20240304
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
    Dim site As Variant 'Carter, 20240304
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
    TheExec.Flow.TestLimit resultVal:=DSP_MTRGR_Gain_eFuse.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
Dim site As Variant 'Carter, 20240304
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
    TheExec.Flow.TestLimit resultVal:=DSP_OffSet_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
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
Dim site As Variant 'Carter, 20240304
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
TheExec.Flow.TestLimit resultVal:=DSP_Gain_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

Call StoreDataAllType(argv(3) & "_" & CStr(TheExec.Flow.var(argv(2)).value), DSP_Gain_Mean)

If TheExec.Flow.var(argv(2)).value = argv(5) Then
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
    TheExec.Flow.TestLimit resultVal:=DSP_Gain_Mean_Final.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
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
    TheExec.Flow.TestLimit resultVal:=BTS_V0, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling ', formatStr:="%.5f"   ''new
       
    Call StoreDataAllType(argv(1), BTS_V0)  '& "_" & CStr(TheExec.Flow.var(argv(2)).Value)
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_OffSet") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    Dim site As Variant 'Carter, 20240304
    For Each site In TheExec.sites.Active
        ARRAY_MTRBTS_OUT1_DEC(0) = DSP_MTRBTS_OUT1
        ARRAY_MTRBTS_OUT2_DEC(0) = DSP_MTRBTS_OUT2
        
        DSP_MTRBTS_OUT1_DEC.data = ARRAY_MTRBTS_OUT1_DEC
        DSP_MTRBTS_OUT2_DEC.data = ARRAY_MTRBTS_OUT2_DEC
        BTS_V1 = DSP_MTRBTS_OUT1_DEC.Subtract(DSP_MTRBTS_OUT2_DEC).divide(2 ^ 16).Element(0)
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=BTS_V1, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.4f" ''try
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
    Dim site As Variant 'Carter, 20240304
    '===================201209 BTS new constant===========================
    Dim A0, a1, a2, a3, Tcp, k0, k1 As Double
        
    'V1
'    A0 = 53.42539937
'    a1 = 106.24914282
'    a2 = -5.64960584
'    a3 = 0.5126488
'    k0 = 133.35292861
'    k1 = 98.97608898
'    Tcp = 24.3
    
'    'V2 High Accuracy mode
    A0 = 53.26973036
    a1 = 104.64623585
    a2 = -8.85459385
    a3 = -0.30123576
    k0 = 131.34112566
    k1 = 97.48290549
    Tcp = 25

    If gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment "a0=" & A0 & ",a1=" & a1 & ",a2=" & a2 & ",a3=" & a3 & ",k0=" & k0 & ",k1=" & k1 & ",Tcp=" & Tcp
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
    Dim c1 As New SiteDouble
    Dim C2 As New SiteDouble
    Dim C3 As New SiteDouble
    Dim C0A As New SiteDouble
    Dim C1A As New SiteDouble
    Dim DSP_C0_Cal_eFuse As New DSPWave
    Dim DSP_C1_Cal_eFuse As New DSPWave
    Dim DSP_C2_Cal_eFuse As New DSPWave
    Dim DSP_C3_Cal_eFuse As New DSPWave
    Dim DSP_C0A_Cal_eFuse As New DSPWave
    Dim DSP_C1A_Cal_eFuse As New DSPWave
    Dim DSP_C0_CAL  As New DSPWave
    Dim DSP_C1_CAL  As New DSPWave
    Dim DSP_C2_CAL  As New DSPWave
    Dim DSP_C3_CAL  As New DSPWave
    Dim DSP_C0A_CAL As New DSPWave
    Dim DSP_C1A_CAL As New DSPWave
    Dim DSP_C0_Cal_Src As New DSPWave
    Dim DSP_C1_Cal_Src As New DSPWave
    Dim DSP_C2_Cal_Src As New DSPWave
    Dim DSP_C3_Cal_Src As New DSPWave
    Dim DSP_C0A_Cal_Src As New DSPWave
    Dim DSP_C1A_Cal_Src As New DSPWave
    
    DSP_C0_Cal_eFuse.CreateConstant 0, 1
    DSP_C1_Cal_eFuse.CreateConstant 0, 1
    DSP_C2_Cal_eFuse.CreateConstant 0, 1
    DSP_C3_Cal_eFuse.CreateConstant 0, 1
    DSP_C0A_Cal_eFuse.CreateConstant 0, 1
    DSP_C1A_Cal_eFuse.CreateConstant 0, 1
    DSP_C0_CAL.CreateConstant 0, 1
    DSP_C1_CAL.CreateConstant 0, 1
    DSP_C2_CAL.CreateConstant 0, 1
    DSP_C3_CAL.CreateConstant 0, 1
    DSP_C0A_CAL.CreateConstant 0, 1
    DSP_C1A_CAL.CreateConstant 0, 1

    
    C0 = A0_CAL(0).Subtract(v0.Multiply(a1).divide(v1))                                    ''A0_CAL(0)- a1 * V0 / V1
    c1 = v1.Invert.Multiply(a1).Subtract(v0.Multiply(2 * a2).divide(v1.Power(2)))          ''a1 / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
    C2 = v1.Power(2).Invert.Multiply(a2).Subtract(v0.Multiply(3 * a3).divide(v1.Power(3))) ''a2 / (BTS_V1 ^ 2) - 3 * a3 * BTS_V0 / (BTS_V1 ^ 3)
    C3 = v1.Power(3).Invert.Multiply(a3)                                                   ''a3 / (BTS_V1 ^ 3)
    C0A = B0_CAL(0).Subtract(B1_CAL(0).Multiply(v0).divide(v1))                            ''B0_CAL(0) - B1_CAL(0) * BTS_V0 / BTS_V1
    C1A = B1_CAL(0).divide(v1).Subtract(v0.Multiply(2 * a2).divide(v1.Power(2)))           ''B1_CAL / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
    '=================================================================
    For Each site In TheExec.sites.Active
        DSP_C0_CAL.Element(0) = C0
        DSP_C1_CAL.Element(0) = c1
        DSP_C2_CAL.Element(0) = C2
        DSP_C3_CAL.Element(0) = C3
        DSP_C0A_CAL.Element(0) = C0A
        DSP_C1A_CAL.Element(0) = C1A
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=x0a(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=x0(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=A0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=B0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=B1_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_C0_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C0_SRC", DSP_C0_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_C1_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C1_SRC", DSP_C1_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_C2_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C2_SRC", DSP_C2_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_C3_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C3_SRC", DSP_C3_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_C0A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C0A_SRC", DSP_C0A_Cal_Src)
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.Flow.TestLimit resultVal:=DSP_C1A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
'    Call StoreDataAllType(Sensor_Num(0) & "_C1A_SRC", DSP_C1A_Cal_Src)
    
    C0 = C0.Multiply(2 ^ 9)
    c1 = c1.Multiply(2 ^ 9)
    C2 = C2.Multiply(2 ^ 9)
    C3 = C3.Multiply(2 ^ 9)
    C0A = C0A.Multiply(2 ^ 9)
    C1A = C1A.Multiply(2 ^ 9)
    For Each site In TheExec.sites.Active
        DSP_C0_CAL.Element(0) = C0
        DSP_C1_CAL.Element(0) = c1
        DSP_C2_CAL.Element(0) = C2
        DSP_C3_CAL.Element(0) = C3
        DSP_C0A_CAL.Element(0) = C0A
        DSP_C1A_CAL.Element(0) = C1A
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
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0A_Cal_eFuse, DSP_C0A_CAL, 20, 0)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1A_Cal_eFuse, DSP_C1A_CAL, 20, 0)
    
    Call HardIP_Dec2Bin(DSP_C0_Cal_Src, DSP_C0_Cal_eFuse, 20) 'DSP_DecToBin
    Call HardIP_Dec2Bin(DSP_C1_Cal_Src, DSP_C1_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C2_Cal_Src, DSP_C2_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C3_Cal_Src, DSP_C3_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C0A_Cal_Src, DSP_C0A_Cal_eFuse, 20)
    Call HardIP_Dec2Bin(DSP_C1A_Cal_Src, DSP_C1A_Cal_eFuse, 20)
    
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites.Active
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C0=" & DSP_C0_CAL.Element(0) & " => " & DSP_C0_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C1=" & DSP_C1_CAL.Element(0) & " => " & DSP_C1_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C2=" & DSP_C2_CAL.Element(0) & " => " & DSP_C2_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C3=" & DSP_C3_CAL.Element(0) & " => " & DSP_C3_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C0A=" & DSP_C0A_CAL.Element(0) & " => " & DSP_C0A_Cal_eFuse.Element(0)
            TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C1A=" & DSP_C1A_CAL.Element(0) & " => " & DSP_C1A_Cal_eFuse.Element(0)
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
    Call StoreDataAllType(Sensor_Num(0) & "_C0A_SRC", DSP_C0A_Cal_Src)
'    TestNameInput = Report_TName_From_Instance("CalcC", "")
'    TheExec.Flow.TestLimit resultVal:=DSP_C1A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    Call StoreDataAllType(Sensor_Num(0) & "_C1A_SRC", DSP_C1A_Cal_Src)
    
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
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoreDataAllType(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).data = Temperature_Array
        Next site
    Next i
    
    For i = 0 To UBound(Temperature_Sensor)
        Sensor_Num = Split(Temperature_Sensor(i), "_")
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
        TheExec.Flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Temperature") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    
    Dim DSP_MTRSNS_Temperature_Average As New DSPWave
    Dim DSP_MTRSNS_Temperature_Maximum As New DSPWave
    Dim DSP_MTRSNS_Temperature_Minimum As New DSPWave
    
    Dim Calc_MAX_FLAG As Boolean: Calc_MAX_FLAG = False 'Add Calculate Flag
    Dim Calc_MIN_FLAG As Boolean: Calc_MIN_FLAG = False 'Add Calculate Flag
    Dim Calc_AVG_FLAG As Boolean: Calc_AVG_FLAG = False 'Add Calculate Flag
    Dim SplitStrAry() As String
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
    DSP_MTRSNS_Temperature_Maximum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    DSP_MTRSNS_Temperature_Minimum.CreateConstant 0, UBound(Temperature_Sensor) + 1
    
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
        
        For Each site In TheExec.sites
            ' Update calc flag judge -- 20220214
            If Calc_AVG_FLAG = True Then DSP_MTRSNS_Temperature_Average(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Average(site).CalcMean, 4)
            If Calc_MAX_FLAG = True Then DSP_MTRSNS_Temperature_Maximum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Maximum(site).CalcMaximumValue, 4)
            If Calc_MIN_FLAG = True Then DSP_MTRSNS_Temperature_Minimum(site).Element(0) = FormatNumber(DSP_MTRSNS_Temperature_Minimum(site).CalcMinimumValue, 4)
        Next site
        
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, , , , , , , tlForceNone)
        
        If UCase(TheExec.Flow.CurrentFlowSheetName) = "FLOW_TMPS" Or UCase(TheExec.Flow.CurrentFlowSheetName) = "FLOW_TMPS_NO_RELAY" Or UCase(TheExec.Flow.CurrentFlowSheetName) = "FLOW_TMPS_TSNS" Then
            Dim TestNameInput_Ary() As String
            Dim glb_MTRRecord As String
            Dim glb_MTRBTSCnt As String
            TestNameInput_Ary = Split(TestNameInput, "_")
            TestNameInput_Ary(8) = glb_MTRRecord & "-" & glb_MTRBTSCnt
            TestNameInput = Join(TestNameInput_Ary, "_")
        End If
        ' Update calc flag judge -- 20220214
        If Calc_AVG_FLAG = True Then TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Average.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Average_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
        If Calc_MAX_FLAG = True Then TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Maximum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Maximum_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
        If Calc_MIN_FLAG = True Then TheExec.Flow.TestLimit resultVal:=DSP_MTRSNS_Temperature_Minimum.Element(0), unit:=unitCustom, Tname:=Replace(TestNameInput, "_X_X_x_", "_X_Temperature_Minimum_", , , vbTextCompare), ForceResults:=tlForceFlow, customUnit:="C"
    End If
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_Metrology_Temperature_Max_Min_Avg") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
        
        SD_V1.Pins(0) = GetStoreDataAllType(Dict_V1)
        SD_V2.Pins(0) = GetStoreDataAllType(Dict_V2)
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
        temp(0) = SD_nts_vref_dig.Pins(0).value
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
    TheExec.Flow.TestLimit resultVal:=SD_nts_vref_dig.Pins(0), Tname:=TestNameInput, ForceResults:=tlForceFlow   'tlForceFlow
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_Metrology_nTS_T1") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
    
End Function


Public Function Calc_MetrologyHSCnTS_STD_Coefficient_N3(argc As Long, argv() As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'Calc_MetrologyHSCnTS_STD_Coefficient_N3(hsc_nts_tp00g_d_0_cp1,hsc_nts_tp00g_vref_dig_0_cp1,nts_vref_dig_G_ft2,25,hsc_nts_tp00g_c0,hsc_nts_tp00g_c1,hsc_nts_tp00g_c2)
    Dim counter As Integer
    Dim str1 As String
    Dim str2 As String
        
    Dim site As Variant
    Dim i As Integer
    Dim Dict_ReadFromEfuse As String
    Dict_ReadFromEfuse = argv(0)
    
    Dim ReadEfuse_Bin As New DSPWave
    Dim ReadEfuse_Dec As New DSPWave
    ReadEfuse_Bin.CreateConstant 0, 21, DspLong
    ReadEfuse_Dec.CreateConstant 0, 1, DspLong
    
    Dim DSPWF_LSB As New DSPWave
    Dim DSPWF_MSB As New DSPWave
    Dim DSPWF_Mid As New DSPWave
    DSPWF_LSB.CreateConstant 0, 16, DspLong
    DSPWF_MSB.CreateConstant 0, 1, DspLong
    DSPWF_Mid.CreateConstant 0, 4, DspLong
    
    Dim x As New SiteLong
    
    
    If TheExec.TesterMode = testModeOffline Then            'for offline run

    'readEfuse=LSB------------------MSB
        str1 = "011110011111100111111"
        str2 = "111110000001101100000"
        
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
    
    For Each site In TheExec.sites
        x = ReadEfuse_Bin.sampleSize - 1
        For i = 0 To x
            If i < 16 Then
                DSPWF_LSB.Element(i) = ReadEfuse_Bin.Element(i)
            ElseIf i = x Then
                DSPWF_MSB.Element(i - x) = ReadEfuse_Bin.Element(i)
            ElseIf i > 15 And i < 21 Then
                DSPWF_Mid.Element(i - 16) = ReadEfuse_Bin.Element(i)
            End If
        Next i
    Next site
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
        y = ReadEfuse2_cp1_Bin.sampleSize - 1
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

    For Each site In TheExec.sites
        BKM_DECODE = gS_BKM_IEDA
        Exit For
    Next site
    
    If MTR_COEFF_INFO_DICT.Exists("TSNS_N3_" & BKM_DECODE) Then
        a0_gold_0 = MTR_COEFF_INFO_DICT("TSNS_N3_" & BKM_DECODE & "_A0")
        a1_gold_0 = MTR_COEFF_INFO_DICT("TSNS_N3_" & BKM_DECODE & "_A1")
        a2_gold_0 = MTR_COEFF_INFO_DICT("TSNS_N3_" & BKM_DECODE & "_A2")
        sBKM = MTR_COEFF_INFO_DICT("TSNS_N3_" & BKM_DECODE & "_BKM")
    Else
        a0_gold_0 = MTR_COEFF_INFO_DICT("TSNS_N3_Default_A0")
        a1_gold_0 = MTR_COEFF_INFO_DICT("TSNS_N3_Default_A1")
        a2_gold_0 = MTR_COEFF_INFO_DICT("TSNS_N3_Default_A2")
        sBKM = MTR_COEFF_INFO_DICT("TSNS_N3_Default_BKM")
    End If
    
    TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM :" & sBKM & ")**********************"
    TheExec.Datalog.WriteComment "  a0_gold_0 = " & a0_gold_0
    TheExec.Datalog.WriteComment "  a1_gold_0 = " & a1_gold_0
    TheExec.Datalog.WriteComment "  a2_gold_0 = " & a2_gold_0
    TheExec.Datalog.WriteComment "*************************************************************************************"

'    a0_gold_0 = 2.08
'    a1_gold_0 = 452.97
'    a2_gold_0 = -135.37
'    TheExec.Datalog.WriteComment "********************** Coefficients for Calibration **********************"
'    TheExec.Datalog.WriteComment "  a0_gold_0 = 2.08"
'    TheExec.Datalog.WriteComment "  a1_gold_0 = 452.97"
'    TheExec.Datalog.WriteComment "  a2_gold_0 = -135.37"
'    TheExec.Datalog.WriteComment "**************************************************************************"
    
     
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
    TheExec.Flow.TestLimit resultVal:=C0_i_2S_C.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.Flow.TestLimit resultVal:=C1_i_2S_C.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.Flow.TestLimit resultVal:=C2_i_2S_C.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyHSCnTS_STD_Coefficient_N3") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

End Function


Public Function Calc_MetrologyHSCnTS_ADV_Coefficient(argc As Long, argv() As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
''Calc_MetrologyHSCnTS_ADV_Coefficient((hsc_nts_tp00g_d_0_cp1,hsc_nts_tp00g_vref_dig_0_cp1,nts_vref_dig_G_ft2,25,hsc_nts_tp00g_c0,hsc_nts_tp00g_c1)

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
    
    Dim Dict_BTS_CP1_Temp As New DSPWave
    Dim Dict_BTS_CP2_Temp As New DSPWave
    Dict_BTS_CP1_Temp.CreateConstant 0, 21, DspLong
    Dict_BTS_CP2_Temp.CreateConstant 0, 21, DspLong
    Dim Dict_12B32X_Temp As New DSPWave
    Dim Dict_nTS_CP1_Temp As New DSPWave
    Dict_12B32X_Temp.CreateConstant 0, 21, DspLong
    Dict_nTS_CP1_Temp.CreateConstant 0, 21, DspLong
    'ReadEfuse_Dec.CreateConstant 0, 1, DspLong
    
    Dim Dict_BTS_CP1 As String
    Dim Dict_BTS_CP2 As String
    Dim Dict_12B32X As String
    Dim Dict_nTS_CP1 As String
    
'    Dim Dict_BTS_CP1_Temp_Str As String
'    Dim Dict_BTS_CP2_Temp_Str As String

    Dim Dict_BTS_CP1_Dec As New DSPWave
    Dim Dict_BTS_CP2_Dec As New DSPWave
    Dim Dict_12B32X_Dec As New DSPWave
    Dim Dict_nTS_CP1_Dec As New DSPWave
    
    Dim Dict_BTS_CP1_Bin As New DSPWave
    Dim Dict_BTS_CP2_Bin As New DSPWave
    Dim Dict_12B32X_Bin As New DSPWave
    Dim Dict_nTS_CP1_Bin As New DSPWave
    
    Dim Dict_BTS_CP1_T As New DSPWave
    Dim Dict_BTS_CP2_T As New DSPWave
    Dim Dict_12B32X_T As New DSPWave
    Dim Dict_nTS_CP1_T As New DSPWave
    Dict_BTS_CP1_T.CreateConstant 0, 21, DspLong
    Dict_BTS_CP2_T.CreateConstant 0, 21, DspLong
    Dict_12B32X_T.CreateConstant 0, 21, DspLong
    Dict_nTS_CP1_T.CreateConstant 0, 21, DspLong

    
    
    Dict_BTS_CP1 = argv(0)
    Dict_BTS_CP2 = argv(1)
    Dict_12B32X = argv(2)
    Dict_nTS_CP1 = argv(3)

'    If TheExec.TesterMode = testModeOffline Then            'for offline run
'    'readEfuse=LSB------------------MSB
'        str1 = "011110011111100"
'        str2 = "111110000001101"
'        str3 = "101110011111100"
'        str4 = "110110000001101"
'        str5 = "111100011111100"
'        str6 = "001110000001101"
'        str7 = "110110011111100"
'        str8 = "111100000001101"
'
'        For Each site In TheExec.sites.Active
'            If site = 0 Then
'                For i = 0 To Len(str1) - 1
'                    Dict_BTS_CP1_Temp.Element(i) = mid(str1, i + 1, 1)
'                    Dict_BTS_CP2_Temp.Element(i) = mid(str3, i + 1, 1)
'                    Dict_12B32X_Temp.Element(i) = mid(str5, i + 1, 1) ''
'                    Dict_nTS_CP1_Temp.Element(i) = mid(str7, i + 1, 1) ''
'                Next i
'            Else
'                For i = 0 To Len(str2) - 1
'                    Dict_BTS_CP1_Temp.Element(i) = mid(str2, i + 1, 1)
'                    Dict_BTS_CP2_Temp.Element(i) = mid(str4, i + 1, 1)
'                    Dict_12B32X_Temp.Element(i) = mid(str6, i + 1, 1) ''
'                    Dict_nTS_CP1_Temp.Element(i) = mid(str8, i + 1, 1) ''
'                Next i
'            End If
'        Next site
'        For Each site In TheExec.sites.Active
'            x = Dict_BTS_CP1_Temp.SampleSize - 1
'            For i = 0 To x
'                Dict_BTS_CP1_T.Element(i) = Dict_BTS_CP1_Temp.Element(i)
'                Dict_BTS_CP2_T.Element(i) = Dict_BTS_CP2_Temp.Element(i)
'                Dict_12B32X_T.Element(i) = Dict_12B32X_Temp.Element(i) ''
'                Dict_nTS_CP1_T.Element(i) = Dict_nTS_CP1_Temp.Element(i) ''
'            Next i
'        Next site
'
'        Call StoreDataAllType(Dict_BTS_CP1, Dict_BTS_CP1_T)
'        Call StoreDataAllType(Dict_BTS_CP2, Dict_BTS_CP2_T)
'        Call StoreDataAllType(Dict_12B32X, Dict_12B32X_T)
'        Call StoreDataAllType(Dict_nTS_CP1, Dict_nTS_CP1_T)
'    End If
    
'    Call Calc_DSP_Dictionary_Process_pa(Dict_BTS_CP1 & "&000000", "Dict_BTS_CP1_Temp_Str")
'    Call Calc_DSP_Dictionary_Process_pa(Dict_BTS_CP2 & "&000000", "Dict_BTS_CP2_Temp_Str")
    
        
    Dict_BTS_CP1_Bin = GetStoreDataAllType(Dict_BTS_CP1)
    Dict_BTS_CP2_Bin = GetStoreDataAllType(Dict_BTS_CP2)
    Dict_12B32X_Bin = GetStoreDataAllType(Dict_12B32X)
    Dict_nTS_CP1_Bin = GetStoreDataAllType(Dict_nTS_CP1)

    Call rundsp.BinToDec(Dict_BTS_CP1_Bin, Dict_BTS_CP1_Dec)
    Call rundsp.BinToDec(Dict_BTS_CP2_Bin, Dict_BTS_CP2_Dec)
    Call rundsp.BinToDec(Dict_12B32X_Bin, Dict_12B32X_Dec)
    Call rundsp.BinToDec(Dict_nTS_CP1_Bin, Dict_nTS_CP1_Dec)


    '------------------
'
    Dim Temp_tp00_c1a As New DSPWave
    Dim Temp_tp00_c1a_temp As New DSPWave
    Dim Temp_tp00_c1a_m As New DSPWave
    Dim Temp_tp00_c1a_2sComplement As New DSPWave
    Dim Temp_tp00_c1a_2sComplement_Bin As New DSPWave
    Dim hsc_nts_tp00_c1a As String
    Dim width As New SiteLong
    width = 21
    
    hsc_nts_tp00_c1a = argv(5)
    Temp_tp00_c1a.CreateConstant 0, 1, DspDouble
    Temp_tp00_c1a_m.CreateConstant 0, 1, DspDouble
    Temp_tp00_c1a_temp.CreateConstant 0, 1, DspDouble
    Temp_tp00_c1a_2sComplement_Bin.CreateConstant 0, 20, DspDouble
    For Each site In TheExec.sites.Active
        Temp_tp00_c1a = Dict_BTS_CP2_Dec.Subtract(Dict_BTS_CP1_Dec)
        Temp_tp00_c1a_m = Dict_12B32X_Dec.Subtract(Dict_nTS_CP1_Dec)
        Temp_tp00_c1a = Temp_tp00_c1a.divide(Temp_tp00_c1a_m)
        Temp_tp00_c1a_temp = Temp_tp00_c1a.Multiply(2 ^ 10) '2^16/2^6=2^10
        Temp_tp00_c1a = Temp_tp00_c1a_temp.Multiply(2 ^ 9) '2^9
    Next site
    Call rundsp.DSP_Convert_2S_Complement(Temp_tp00_c1a, width, Temp_tp00_c1a_2sComplement)
    Call rundsp.DSPWaveDecToBinary(Temp_tp00_c1a_2sComplement, 20, Temp_tp00_c1a_2sComplement_Bin)
    Call StoreDataAllType(hsc_nts_tp00_c1a, Temp_tp00_c1a_2sComplement_Bin)
    
    Dim Temp_tp00_c0a As New DSPWave
    Dim Temp_tp00_c0a_m As New DSPWave
    Dim Temp_tp00_c0a_2sComplement As New DSPWave
    Dim Temp_tp00_c0a_2sComplement_Bin As New DSPWave
    Dim hsc_nts_tp00_c0a As String
    
    hsc_nts_tp00_c0a = argv(4)
    Temp_tp00_c0a.CreateConstant 0, 1, DspDouble
    Temp_tp00_c0a_2sComplement_Bin.CreateConstant 0, 20, DspDouble
    
    For Each site In TheExec.sites.Active
        Temp_tp00_c0a = Dict_BTS_CP1_Dec.divide(2 ^ 6) '2^6
        Temp_tp00_c0a = Temp_tp00_c0a.Subtract(Temp_tp00_c1a_temp.Multiply(Dict_nTS_CP1_Dec).divide(2 ^ 16)) '2^16
        Temp_tp00_c0a = Temp_tp00_c0a.Multiply(2 ^ 9) '2^9
    Next site
    Call rundsp.DSP_Convert_2S_Complement(Temp_tp00_c0a, width, Temp_tp00_c0a_2sComplement)
    Call rundsp.DSPWaveDecToBinary(Temp_tp00_c0a_2sComplement, 20, Temp_tp00_c0a_2sComplement_Bin)
    Call StoreDataAllType(hsc_nts_tp00_c0a, Temp_tp00_c0a_2sComplement_Bin)
    
    For Each site In TheExec.sites.Active
    
        TheExec.Datalog.WriteComment "site_" & site & "_" & argv(0) & " : " & Dict_BTS_CP1_Dec(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & argv(1) & " : " & Dict_BTS_CP2_Dec(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & argv(2) & " : " & Dict_12B32X_Dec(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & argv(3) & " : " & Dict_nTS_CP1_Dec(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & hsc_nts_tp00_c0a & " : " & Temp_tp00_c0a_2sComplement(site).Element(0)
        TheExec.Datalog.WriteComment "site_" & site & "_" & hsc_nts_tp00_c1a & " : " & Temp_tp00_c1a_2sComplement(site).Element(0)

    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.Flow.TestLimit resultVal:=Temp_tp00_c0a_2sComplement.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.Flow.TestLimit resultVal:=Temp_tp00_c1a_2sComplement.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow

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
    Dim A0, a1, a2, a3, Tcp, k0, k1 As Double
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
    
    Dim site As Variant 'Carter, 20240304
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
            A0 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_A0"))
            k1 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_K1"))
            k0 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_K0"))
            Tcp = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_TCP"))
            sBKM = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_" & BKM_DECODE & "_BKM"))
        Else
            a3 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_A3"))
            a2 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_A2"))
            a1 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_A1"))
            A0 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_A0"))
            k1 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_K1"))
            k0 = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_K0"))
            Tcp = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_TCP"))
            sBKM = MTR_COEFF_INFO_DICT(UCase("BTS_N3E_Default_BKM"))
        End If
    
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM :" & sBKM & ")**********************"
        TheExec.Datalog.WriteComment "  a3 = " & a3
        TheExec.Datalog.WriteComment "  a2 = " & a2
        TheExec.Datalog.WriteComment "  a1 = " & a1
        TheExec.Datalog.WriteComment "  a0 = " & A0
        TheExec.Datalog.WriteComment "  k1 = " & k1
        TheExec.Datalog.WriteComment "  k0 = " & k0
        TheExec.Datalog.WriteComment "  Tcp = " & Tcp
        TheExec.Datalog.WriteComment "*************************************************************************************"
    
    
        If gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "a0=" & A0 & ",a1=" & a1 & ",a2=" & a2 & ",a3=" & a3 & ",k0=" & k0 & ",k1=" & k1 & ",Tcp=" & Tcp
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
        Dim c1 As New SiteDouble
        Dim C2 As New SiteDouble
        Dim C3 As New SiteDouble
        Dim C0A As New SiteDouble
        Dim C1A As New SiteDouble
        Dim DSP_C0_Cal_eFuse As New DSPWave
        Dim DSP_C1_Cal_eFuse As New DSPWave
        Dim DSP_C2_Cal_eFuse As New DSPWave
        Dim DSP_C3_Cal_eFuse As New DSPWave
        Dim DSP_C0A_Cal_eFuse As New DSPWave
        Dim DSP_C1A_Cal_eFuse As New DSPWave
        Dim DSP_C0_CAL  As New DSPWave
        Dim DSP_C1_CAL  As New DSPWave
        Dim DSP_C2_CAL  As New DSPWave
        Dim DSP_C3_CAL  As New DSPWave
        Dim DSP_C0A_CAL As New DSPWave
        Dim DSP_C1A_CAL As New DSPWave
        Dim DSP_C0_Cal_Src As New DSPWave
        Dim DSP_C1_Cal_Src As New DSPWave
        Dim DSP_C2_Cal_Src As New DSPWave
        Dim DSP_C3_Cal_Src As New DSPWave
        Dim DSP_C0A_Cal_Src As New DSPWave
        Dim DSP_C1A_Cal_Src As New DSPWave
        
        
        Set DSP_C0_Cal_eFuse = Nothing
        Set DSP_C1_Cal_eFuse = Nothing
        Set DSP_C2_Cal_eFuse = Nothing
        Set DSP_C3_Cal_eFuse = Nothing
        Set DSP_C0A_Cal_eFuse = Nothing
        Set DSP_C1A_Cal_eFuse = Nothing
        Set DSP_C0_CAL = Nothing
        Set DSP_C1_CAL = Nothing
        Set DSP_C2_CAL = Nothing
        Set DSP_C3_CAL = Nothing
        Set DSP_C0A_CAL = Nothing
        Set DSP_C1A_CAL = Nothing
        
        'If k = 0 Then
        DSP_C0_Cal_eFuse.CreateConstant 0, 1
        DSP_C1_Cal_eFuse.CreateConstant 0, 1
        DSP_C2_Cal_eFuse.CreateConstant 0, 1
        DSP_C3_Cal_eFuse.CreateConstant 0, 1
        DSP_C0A_Cal_eFuse.CreateConstant 0, 1
        DSP_C1A_Cal_eFuse.CreateConstant 0, 1
        DSP_C0_CAL.CreateConstant 0, 1
        DSP_C1_CAL.CreateConstant 0, 1
        DSP_C2_CAL.CreateConstant 0, 1
        DSP_C3_CAL.CreateConstant 0, 1
        DSP_C0A_CAL.CreateConstant 0, 1
        DSP_C1A_CAL.CreateConstant 0, 1
        'End If
        
        C0 = A0_CAL(0).Subtract(v0.Multiply(a1).divide(v1))                                    ''A0_CAL(0)- a1 * V0 / V1
        c1 = v1.Invert.Multiply(a1).Subtract(v0.Multiply(2 * a2).divide(v1.Power(2)))          ''a1 / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
        C2 = v1.Power(2).Invert.Multiply(a2).Subtract(v0.Multiply(3 * a3).divide(v1.Power(3))) ''a2 / (BTS_V1 ^ 2) - 3 * a3 * BTS_V0 / (BTS_V1 ^ 3)
        C3 = v1.Power(3).Invert.Multiply(a3)                                                   ''a3 / (BTS_V1 ^ 3)
        C0A = B0_CAL(0).Subtract(B1_CAL(0).Multiply(v0).divide(v1))                            ''B0_CAL(0) - B1_CAL(0) * BTS_V0 / BTS_V1
        C1A = B1_CAL(0).divide(v1).Subtract(v0.Multiply(2 * a2).divide(v1.Power(2)))           ''B1_CAL / BTS_V1 - 2 * a2 * BTS_V0 / (BTS_V1 ^ 2)
        '=================================================================
        For Each site In TheExec.sites.Active
            DSP_C0_CAL.Element(0) = C0
            DSP_C1_CAL.Element(0) = c1
            DSP_C2_CAL.Element(0) = C2
            DSP_C3_CAL.Element(0) = C3
            DSP_C0A_CAL.Element(0) = C0A
            DSP_C1A_CAL.Element(0) = C1A
            
            All_sensor_C0.Element(k) = C0
        Next site
        If gl_Disable_HIP_debug_log = False Then
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=x0a(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=x0(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=A0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=B0_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=B1_CAL(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=DSP_C0_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=DSP_C1_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=DSP_C2_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=DSP_C3_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=DSP_C0A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=DSP_C1A_CAL.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        Else
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 11
        End If
        
        C0 = C0.Multiply(2 ^ 9)
        c1 = c1.Multiply(2 ^ 9)
        C2 = C2.Multiply(2 ^ 9)
        C3 = C3.Multiply(2 ^ 9)
        C0A = C0A.Multiply(2 ^ 9)
        C1A = C1A.Multiply(2 ^ 9)
        For Each site In TheExec.sites.Active
            DSP_C0_CAL.Element(0) = C0
            DSP_C1_CAL.Element(0) = c1
            DSP_C2_CAL.Element(0) = C2
            DSP_C3_CAL.Element(0) = C3
            DSP_C0A_CAL.Element(0) = C0A
            DSP_C1A_CAL.Element(0) = C1A
        Next site
        
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0_Cal_eFuse, DSP_C0_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1_Cal_eFuse, DSP_C1_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C2_Cal_eFuse, DSP_C2_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C3_Cal_eFuse, DSP_C3_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C0A_Cal_eFuse, DSP_C0A_CAL, 20, 0)
        Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_C1A_Cal_eFuse, DSP_C1A_CAL, 20, 0)
        
        
        Set DSP_C0_Cal_Src = Nothing
        Set DSP_C1_Cal_Src = Nothing
        Set DSP_C2_Cal_Src = Nothing
        Set DSP_C3_Cal_Src = Nothing
        Set DSP_C0A_Cal_Src = Nothing
        Set DSP_C1A_Cal_Src = Nothing
        
        Call HardIP_Dec2Bin(DSP_C0_Cal_Src, DSP_C0_Cal_eFuse, 20) 'DSP_DecToBin
        Call HardIP_Dec2Bin(DSP_C1_Cal_Src, DSP_C1_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C2_Cal_Src, DSP_C2_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C3_Cal_Src, DSP_C3_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C0A_Cal_Src, DSP_C0A_Cal_eFuse, 20)
        Call HardIP_Dec2Bin(DSP_C1A_Cal_Src, DSP_C1A_Cal_eFuse, 20)
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C0=" & DSP_C0_CAL.Element(0) & " => " & DSP_C0_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C1=" & DSP_C1_CAL.Element(0) & " => " & DSP_C1_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C2=" & DSP_C2_CAL.Element(0) & " => " & DSP_C2_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C3=" & DSP_C3_CAL.Element(0) & " => " & DSP_C3_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C0A=" & DSP_C0A_CAL.Element(0) & " => " & DSP_C0A_Cal_eFuse.Element(0)
                TheExec.Datalog.WriteComment "site_" & site & "_" & Sensor_Num(0) & "_C1A=" & DSP_C1A_CAL.Element(0) & " => " & DSP_C1A_Cal_eFuse.Element(0)
            Next site
        End If
        
   
        Call AddStoredCaptureData(Sensor_Num(0) & "_C0_SRC", DSP_C0_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C1_SRC", DSP_C1_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C2_SRC", DSP_C2_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C3_SRC", DSP_C3_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C0A_SRC", DSP_C0A_Cal_Src)
        Call AddStoredCaptureData(Sensor_Num(0) & "_C1A_SRC", DSP_C1A_Cal_Src)
        
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
        
        If C0_max_value_index_dsp.sampleSize = 0 Then
            TheExec.Datalog.WriteComment "[ERROR] Did't find site" & site & "max Coefficient C0 !!!!"
            Exit For
        Else
            For i = 0 To C0_max_value_index_dsp.sampleSize - 1
                If i = 0 Then
                    temp_max_sensor_name = Split(Split(argv(C0_max_value_index_dsp(site).data(i)), "+")(0), "_")(0)
                Else
                    temp_max_sensor_name = temp_max_sensor_name & "&" & Split(Split(argv(C0_max_value_index_dsp(site).data(i)), "+")(0), "_")(0)
                End If
            Next i
        End If
        
        If C0_min_value_index_dsp.sampleSize = 0 Then
            TheExec.Datalog.WriteComment "[ERROR] Did't find site" & site & "min Coefficient C0!!!!"
            Exit For
        Else
            For i = 0 To C0_min_value_index_dsp.sampleSize - 1
                If i = 0 Then
                    temp_min_sensor_name = Split(Split(argv(C0_min_value_index_dsp(site).data(i)), "+")(0), "_")(0)
                Else
                    temp_min_sensor_name = temp_min_sensor_name & "&" & Split(Split(argv(C0_min_value_index_dsp(site).data(i)), "+")(0), "_")(0)
                End If
            Next i
        End If
        
        TestNameInput = Report_TName_From_Instance("CalcC", temp_max_sensor_name)
        TheExec.Flow.TestLimit resultVal:=C0_max_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        TestNameInput = Report_TName_From_Instance("CalcC", temp_min_sensor_name)
        TheExec.Flow.TestLimit resultVal:=C0_min_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        TestNameInput = Report_TName_From_Instance("CalcC", "")
        TheExec.Flow.TestLimit resultVal:=C0_avg_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        
        If gl_Disable_HIP_debug_log = False Then
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=C0_max2avg_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=C0_min2avg_value(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 5
        Else
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 3
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
            TheExec.Flow.TestLimit resultVal:=DSP_Temperature.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next i
    Else
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + UBound(Temperature_Sensor) + 1
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
        
        If max_temperature_index_dsp.sampleSize = 0 Then
            TheExec.Datalog.WriteComment "[ERROR] Did't find site" & site & "max temperature!!!!"
            Exit For
        Else
            For i = 0 To max_temperature_index_dsp.sampleSize - 1
                If i = 0 Then
                    temp_max_sensor_name = Split(Temperature_Sensor(max_temperature_index_dsp(site).data(i)), "_")(0)
                Else
                    temp_max_sensor_name = temp_max_sensor_name & "&" & Split(Temperature_Sensor(max_temperature_index_dsp(site).data(i)), "_")(0)
                End If
            Next i
        End If
        
        If max_temperature_index_dsp.sampleSize = 0 Then
            TheExec.Datalog.WriteComment "[ERROR] Did't find site" & site & "min temperature!!!!"
            Exit For
        Else
            For i = 0 To min_temperature_index_dsp.sampleSize - 1
                If i = 0 Then
                    temp_min_sensor_name = Split(Temperature_Sensor(min_temperature_index_dsp(site).data(i)), "_")(0)
                Else
                    temp_min_sensor_name = temp_min_sensor_name & "&" & Split(Temperature_Sensor(min_temperature_index_dsp(site).data(i)), "_")(0)
                End If
            Next i
        End If

        
        TestNameInput = Report_TName_From_Instance("CalcC", temp_max_sensor_name)
        TheExec.Flow.TestLimit resultVal:=Max_sensor_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        TestNameInput = Report_TName_From_Instance("CalcC", temp_min_sensor_name)
        TheExec.Flow.TestLimit resultVal:=Min_sensor_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        TestNameInput = Report_TName_From_Instance("CalcC", "")
        TheExec.Flow.TestLimit resultVal:=Avg_sensor_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
        
        If gl_Disable_HIP_debug_log = False Then
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=Max2Avg_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            TestNameInput = Report_TName_From_Instance("CalcC", "")
            TheExec.Flow.TestLimit resultVal:=Min2Avg_temperature(site), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 5
        Else
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 3
        End If
        
    Next site
    
Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Metrology_AP", "Calc_MetrologyBTS_Temperature_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Reg_Assign_Processing_MTRICTS_BPBN(ByRef DigSrc_Equation As String, ByRef DigSrc_Assignment As String, ByRef CUS_Str_DigCapData As String, ByRef Calc_Eqn As String) As Long
    Dim i As Long
    Dim SplitInstance() As String
    Dim ReCombineAssign() As String
    Dim CombineAssignTemp As String
    On Error GoTo errHandler:
    
    If UCase(DigSrc_Equation) Like "*REG_ASSIGN*" Then
        DigSrc_Equation = RegAssign_PatBurstReConstruction(DigSrc_Equation, "Equation_")

    End If
        
    If UCase(DigSrc_Assignment) Like "*REG_ASSIGN*" Then
        DigSrc_Assignment = RegAssign_PatBurstReConstruction(DigSrc_Assignment, "Assignment_")

    End If
    
    If UCase(CUS_Str_DigCapData) Like "*REG_ASSIGN*" Then
        SplitInstance = Split(CUS_Str_DigCapData, ":")
        CUS_Str_DigCapData = Replace(CUS_Str_DigCapData, SplitInstance(0) & ":" & SplitInstance(1), RegDict(LCase("DigCapData_" & SplitInstance(1))))
    End If
    
    If UCase(Calc_Eqn) Like "*REG_ASSIGN*" Then
        SplitInstance = Split(Calc_Eqn, ":")
        Calc_Eqn = RegDict(LCase("Calc_Eqn_" & SplitInstance(1)))
    End If
        

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP", "Reg_Assign_Processing_MTRICTS_BPBN")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ICTS_DAP_PN_TRIM_DigCapDataProcessByDSP(Pat As String, srcPin As PinList, DAC_bp() As SiteLong, DAC_bn() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, TestSequenceNumber As Long, ByRef Rtn_DigCapData() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName_Array_BP() As String, TrimStoreName_Array_BN() As String, DigCap_Pin As PinList, Cap_Trimwidth As Long, DigCap_Sample_Size As Long, CUS_Str_MainProgram As String, CUS_Str_DigCapData As String, width_Wf As DSPWave, width_Wf_2S As DSPWave, ByRef OutWf_2s As DSPWave)
    Dim srcWave_bp() As New DSPWave: ReDim srcWave_bp(UBound(DAC_bp))
    Dim srcWave_bn() As New DSPWave: ReDim srcWave_bn(UBound(DAC_bn))
    Dim capWave_bp() As New DSPWave: ReDim capWave_bp(UBound(DAC_bp))
    Dim capWave_bn() As New DSPWave: ReDim capWave_bn(UBound(DAC_bn))
    Dim site As Variant
    Dim InDSPWave As New DSPWave
    Dim i As Integer, j As Integer
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array_bp() As Long: ReDim srcwave_array_bp(TrimCodeSize - 1)
    Dim srcwave_array_bn() As Long: ReDim srcwave_array_bn(TrimCodeSize - 1)
    Dim OutDspWave As New DSPWave
    Dim OutBinWf As New DSPWave
    Dim TestOutDspWave As New DSPWave
    Dim OutWf As New DSPWave
    Dim DecWave_bp() As New DSPWave: ReDim DecWave_bp(UBound(DAC_bp))
    Dim DecWave_bn() As New DSPWave: ReDim DecWave_bn(UBound(DAC_bn))
'    Dim trim_Wf As New DSPWave: trim_Wf.CreateConstant 0, (TestSequenceNumber / 2 - 1)
    On Error GoTo errHandler
    
    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True

    For i = 0 To UBound(DAC_bp)
    srcWave_bp(i).CreateConstant 0, TrimCodeSize, DspLong
    srcWave_bn(i).CreateConstant 0, TrimCodeSize, DspLong
    DecWave_bp(i).CreateConstant 0, 1, DspLong
    DecWave_bn(i).CreateConstant 0, 1, DspLong
    


        For Each site In TheExec.sites
        DecWave_bp(i).Element(0) = DAC_bp(i)
        DecWave_bn(i).Element(0) = DAC_bn(i)
            For j = 0 To TrimCodeSize - 1
                If j = 0 Then
                    srcwave_array_bp(j) = DAC_bp(i) And 1
                    srcwave_array_bn(j) = DAC_bn(i) And 1
                Else
                    srcwave_array_bp(j) = (DAC_bp(i) And (2 ^ j)) \ (2 ^ j)
                    srcwave_array_bn(j) = (DAC_bn(i) And (2 ^ j)) \ (2 ^ j)
                End If
            Next j
         
        srcWave_bp(i).data = srcwave_array_bp
        srcWave_bn(i).data = srcwave_array_bn
        Next site
        Call StoreDataAllType(TrimStoreName_Array_BP(i), srcWave_bp(i))
        Call StoreDataAllType(TrimStoreName_Array_BP(i) & "_para", DecWave_bp(i))
        Call StoreDataAllType(TrimStoreName_Array_BN(i), srcWave_bn(i))
        Call StoreDataAllType(TrimStoreName_Array_BN(i) & "_para", DecWave_bn(i))
        Next i
        
    Call GeneralDigSrcSettingWithBurst(LCase(Pat), srcPin, InDSPWave)
  


    TheHdw.Patterns(Pat).Load

    Set OutDspWave = Nothing
    Set TestOutDspWave = Nothing
    Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    Call TheHdw.Patterns(Pat).test(pfNever, 0)
    Call CreateSimulateDataDSPWave(OutDspWave, Instance_Data.DigCap_Sample_Size, Cap_Trimwidth)
'    TestOutDspWave.CreateConstant 0, DigCap_Sample_Size, DspLong
    For Each site In TheExec.sites
    TestOutDspWave = OutDspWave.Copy
    Next site
    rundsp.Split_Dspwave OutDspWave, width_Wf, OutBinWf
'    theexec.Flow.TestLimit OutDspWave.Element(0), , ForceResults:=tlForceNone
'    rundsp.Split_2SComplementDSPWave_To_SignDec OutDspWave, width_Wf_2S, OutWf_2s
    rundsp.Split_2SComplementDSPWave_To_SignDec TestOutDspWave, width_Wf_2S, OutWf_2s
    

    For i = 0 To TestSequenceNumber - 1
            Rtn_DigCapData(i).Pins(0).value = OutBinWf.Element(i)
            Call DebugPrintFunc_PPMU("")
            For Each site In TheExec.sites
                Res(i) = Rtn_DigCapData(i).Pins(0).value
                If gl_Disable_HIP_debug_log = False Then
                    If i Mod 2 = 0 Then
                        TheExec.Datalog.WriteComment "Site " & site & ", Source Code DAC_BP:" & DAC_bp(Floor(i / 2)) & ", DSSC Captured Register : " & Rtn_DigCapData(i).Pins(0) & ", DSSC Captured Data = " & Res(i)
                    Else
                        TheExec.Datalog.WriteComment "Site " & site & ", Source Code DAC_BP:" & DAC_bp(Floor(i / 2)) & ", DSSC Captured Register : " & Rtn_DigCapData(i).Pins(0) & ", DSSC Captured Data = " & Res(i)
                    End If
                End If
            Next site
    Next i


    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in ICTS_DAP_PN_TRIM_DigCapDataProcessByDSP"
    If AbortTest Then Exit Function Else Resume Next
End Function

