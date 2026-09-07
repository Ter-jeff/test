Attribute VB_Name = "VBT_LIB_DC_IDS"
#Const isUFP = True
Option Explicit
'Revision History:
'V0.0 initial bring up

Private Const VBT_LIB_DC_IDS = "VBT_LIB_DC_IDS"

'IDS meas pinlistdata
Public IDS_meas As New PinListData

'Delta IDS for iEDA
Public gS_delta_IDS_pcpu As New SiteVariant
Public gS_delta_IDS_ecpu As New SiteVariant
Public gS_delta_IDS_gpu As New SiteVariant
Public gS_delta_IDS_dcs_ddr As New SiteVariant
Public gS_delta_IDS_cpu_sram As New SiteVariant
Public gS_delta_IDS_ave As New SiteVariant

'0826_SPI_IDS measure
Private HexVS_data_SPI_IDS As New PinListData
Private UVS_Hi_data_SPI_IDS As New PinListData
Private UVS_Lo_data_SPI_IDS As New PinListData

'==================================20180928  global variable for eFuse
Public All_Power_data_IDS_GB As New PinListData
''Start, IDS INFO for Efuse and BinCut - Carter, 20190829
Public ids_info_ary() As IDS_INFO

Public gl_IDS_INFO_Dic As New Scripting.Dictionary
Public gl_Dic_IDS_fuse_name_AllCore_25C As New Scripting.Dictionary
Public gl_Dic_IDS_fuse_name_9Core_25C As New Scripting.Dictionary
Public gl_Dic_IDS_fuse_name_AllCore_85C As New Scripting.Dictionary
Public gl_Dic_IDS_fuse_name_9Core_105C As New Scripting.Dictionary
Public gl_Dic_IDS_pin_name As New Scripting.Dictionary

Public gb_IDSLimit_Speical_Handle As Boolean
Public gDict_IDSLimit_Special_Handle As New Scripting.Dictionary

Type IDS_INFO
    pin As String
    Pat As String
    loLimit As String
    hiLimit As String
    MeasureValue As New SiteDouble
End Type
'' old Calc_IDS_Sum
''Public Function Calc_IDS_Sum(argc As Integer, argv() As String) As Long
''
''Dim IDS_String As String
''Dim IDS_Array() As String
''Dim ids_name As Variant
''Dim Dict_Name As String
''Dim Tname_LimitIndex As String
''
''Dim IDSTname As String
''
''Dim Sum_Result As New SiteDouble
''
''
''IDS_String = argv(0)
''
''If UBound(argv) > 0 Then Dict_Name = argv(1)
''
''IDS_Array = Split(IDS_String, "+")
''
''For Each ids_name In IDS_Array
''
''    Sum_Result = Sum_Result.Add(All_Power_data_IDS_GB.Pins(ids_name))
''
''Next ids_name
''
''
''Call GetFlowTName
''Tname_LimitIndex = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
''
''Call AddStoredMeasurement(Tname_LimitIndex, Sum_Result)
''
''IDSTname = Report_TName_From_Instance(CalcI, "X")
''
''TheExec.Flow.TestLimit Sum_Result, ForceResults:=tlForceFlow, Tname:=IDSTname
''
''End Function

''End, IDS INFO for Efuse and BinCut - Carter, 20190829

' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
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


' [20230425][All][Cylde] Restore IFold limit avoid IFold timeout alarm when process Shmoo
' [20230907][All][Oliver] modify for autorange get wrong step accuracy
' [20231003][All][Tank] modify after Chihome review
' [20231016][All][Tank] Use ApplyLevelsTiming to set iFold limit
' [20231124][All][Tank] Enhance Current Profile use simulation value
' [20231124][All][Tank] Fix AutoRange Exit loop
' [20231124][T-Don][Willian] Fix Hilimit not set first currentrange cannot autorange
' [20231228][T-Bra][Tank] Modify FixCurrentRange use rule
' [20240110][All][Tank] Modify fixcurrentrange use function
Public Function DCVS_IDS_main_auto_range_and_measure(CorePower_Pin As String, _
                                                     OtherPower_Pin As String, _
                                                     Power_data As PinListData, _
                                                     repeat_count As Long, _
                                                     FlowLimitForInitIRange As Boolean, _
                                            Optional Search_Step As String, _
                                            Optional debug_print_pins As String, _
                                            Optional AutoRange_Pin As String, _
                                            Optional FixCurrentRange As String, _
                                            Optional isUse_Product_Identifier As Boolean = False, _
                                            Optional nSpecific_Product_Identifier As String = "999")    'FixCurrentRange 20220627
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim i As Long, j As Long, All_Power_Pin As String
    Dim site As Variant, pin As Variant, val As Double
    Dim k As Long
    Dim Tname As String
    Dim VMain As Double
    Dim p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim CorePowerPin_Ary() As String, CorePowerPin_Cnt As Long
    Dim OtherPowerPin_Ary() As String, OtherPowerPin_Cnt As Long
    Dim Hilimit_IDS As Double
    Dim Lolimit_IDS As Double
    Dim n As Long
    
    Dim funcName As String:: funcName = "DCVS_IDS_main_auto_range_and_measure"
    
    ' Get the limits info
    Dim FlowLimitsInfo As IFlowLimitsInfo
    
    ' if no Use-Limits on this test, FlowLimitsInfo is nothing

    Dim Val_Hi() As String
    Dim Val_Lo() As String
    

''    If (FlowLimitForInitIRange = False) Then
''        InitIRange_CreateDic_DCVS HexInitIRange1A_Pins, HexInitIRange100mA_Pins, _
''                            UVSInitIRange800mA_Pins, UVSInitIRange200mA_Pins, UVSInitIRange20mA_Pins, UVSInitIRange2mA_Pins, UVSInitIRange200uA_Pins, _
''                            VSMInitIRange1A_Pins, VSMInitIRange11A_Pins, VSMInitIRange21A_Pins, VSMInitIRange51A_Pins, VSMInitIRange81A_Pins
'''                            UVI80InitIRange1A_Pins , UVI80InitIRange200mA_Pins, UVI80InitIRange20mA_Pins, UVI80InitIRange2mA_Pins, UVI80InitIRange200uA_Pins
''
''    End If

    Dim SattleTime As Double
    Dim WaitTime As Double
    Dim p_hexvs As String
    Dim p_uvs As String
    Dim A_HexVS() As String
    Dim A_UVS() As String
    Dim HexVS_Power_data As New PinListData
    Dim UVS_Power_data As New PinListData
    
    Dim p_vsm As String
    Dim A_VSM() As String
    Dim VSM_Power_data As New PinListData
    
    '''-----------------UFP-----------------
    Dim p_ufp As String
    Dim A_UFP() As String
    Dim UFP_Power_data As New PinListData
    '''-----------------UFP-----------------

    Dim PinVal As New PinData
    Dim DropRngSite As New SiteBoolean
    Dim AutoRangePin_Ary() As String
    Dim AutoRangePin_Cnt As Long
    Dim Power_data_temp As New PinListData
    Dim siteInfo As Variant
    Dim FixCurrentSplitRange() As String 'FixCurrentRange 20220627
    
    Dim s_FixedCurrent As String
    Dim s_FixedCurrent_Pin  As String
    Dim s_FixedCurrent_Ary() As String
    Dim s_FixedCurrent_PinAry()  As String
    Dim dict_FixedCurrent As New Scripting.Dictionary
    
    Dim ids_range_ary() As AutoRange_Info
    Dim Stop_Step, StepNo As Integer
    Dim p_vsm_autorRange As String
    Dim p_hexvs_autorRange As String
    Dim p_uvs256_autorRange As String
    Dim A_vsm_autorRange() As String
    Dim A_hexvs_autorRange() As String
    Dim A_uvs256_autorRange() As String
    Dim VSM_Power_data_AutoRange As New PinListData
    Dim HexVS_Power_data_AutoRange As New PinListData
    Dim UVS256_Power_data_AutoRange As New PinListData
    '''-----------------UFP-----------------
    Dim p_ufp_autorRange As String
    Dim A_UFP_autorange() As String
    Dim UFP_Power_data_AutorRange As New PinListData
    '''-----------------UFP-----------------
    Dim Flag_PerPin_autorange As Boolean
    Dim Flag_AllPin_autorange As Boolean
    Dim nBinCutNumber As New SiteLong
    Dim field_normal As eFuseBdfField
    Dim nTempBinCutNum As New SiteLong
    Dim s_InstrumentName As String
    Dim dic_MeasurePowerPinIndex As New Dictionary
    Dim n_Index As Integer
    Dim s_ErrorMsg As String
    Dim d_SimulationValue As Double
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    ' if no Use-Limits on this test, FlowLimitsInfo is nothing
    If FlowLimitsInfo Is Nothing Then
        If isDebugMode Then TheExec.AddOutput "Could not get the limits info", vbRed, True
        Exit Function
    End If
    
    FlowLimitsInfo.GetHighLimits Val_Hi
    FlowLimitsInfo.GetLowLimits Val_Lo
    
    '20160113: debug with bruce to fix NC issue -- start
    TheExec.DataManager.DecomposePinList CorePower_Pin, CorePowerPin_Ary, CorePowerPin_Cnt
    TheExec.DataManager.DecomposePinList OtherPower_Pin, OtherPowerPin_Ary, OtherPowerPin_Cnt
    
    For i = 0 To CorePowerPin_Cnt - 1
        If gl_GetInstrument_Dic.Exists(LCase(CorePowerPin_Ary(i))) Then All_Power_Pin = CombineStringList(All_Power_Pin, CorePowerPin_Ary(i))
    Next i
    
    For i = 0 To OtherPowerPin_Cnt - 1
        If gl_GetInstrument_Dic.Exists(LCase(OtherPowerPin_Ary(i))) Then All_Power_Pin = CombineStringList(All_Power_Pin, OtherPowerPin_Ary(i))
    Next i

    '20160113: debug with bruce to fix NC issue -- end
    
    Pin_Ary = Split(All_Power_Pin, ",")
    
    Call FixCurrentRange_StrToDic(FixCurrentRange, dict_FixedCurrent)
    
'    All_Power_Pin = CorePower_Pin & "," & OtherPower_Pin
'    TheExec.DataManager.DecomposePinList All_Power_Pin, Pin_Ary, Pin_Cnt
    WaitTime = 100 * us
    
''**** Start - Get PowerPin Info ****
    
    ReDim ids_range_ary(UBound(Pin_Ary))
    
    For i = 0 To UBound(Pin_Ary)
        If dic_MeasurePowerPinIndex.Exists(LCase(Pin_Ary(i))) Then
            s_ErrorMsg = " CorePower and OtherPower have the same pin : " & Pin_Ary(i) & " !!"
            Call Print_Error_Message(Warning_Info, "VBT_LIB_DC_IDS", "DCVS_IDS_main_auto_range_and_measure", s_ErrorMsg)
        Else
            dic_MeasurePowerPinIndex.Add LCase(Pin_Ary(i)), i
            ids_range_ary(i) = PowerPin_range_ary(gl_dicPowerPinIndex(LCase(Pin_Ary(i))))
        End If
    Next i
    
''**** End - Get PowerPin Info ****
        
''**** Start - Set init IRange ****
    For i = 0 To UBound(Pin_Ary)

        ids_range_ary(i).Init_CurrentRange = thehdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.value
               
        ids_range_ary(i).hiLimit = Compare_ForceVal_BV(Pin_Ary(i), Val_Hi(i), True)  'if pins without limit use bincut spec IDSmax as current range
        ids_range_ary(i).loLimit = Compare_ForceVal_BV(Pin_Ary(i), Val_Lo(i), False) 'if pins without limit use bincut spec 0.1IDSmax as current range
        
        s_InstrumentName = UCase(ids_range_ary(i).ChanMapType)
        
        Select Case s_InstrumentName
            Case glbConstIns_HEXVS: p_hexvs = CombineStringList(p_hexvs, Pin_Ary(i))
            Case glbConstIns_VHDVS: p_uvs = CombineStringList(p_uvs, Pin_Ary(i))
            Case glbConstIns_VSM: p_vsm = CombineStringList(p_vsm, Pin_Ary(i))
            Case glbConstIns_VS5A: p_ufp = CombineStringList(p_ufp, Pin_Ary(i))
            Case glbConstIns_VS800MA: p_ufp = CombineStringList(p_ufp, Pin_Ary(i))
        End Select
        
        If dict_FixedCurrent.Exists(LCase(Pin_Ary(i))) Then
            val = CDbl(dict_FixedCurrent(LCase(Pin_Ary(i))))
        ElseIf (FlowLimitForInitIRange = False) Then
            val = thehdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.level.value    'InitIRange_Setup ids_range_ary(i), WaitTime, SattleTime
        Else
            If ids_range_ary(i).hiLimit = 0 Then
                val = ids_range_ary(i).Init_CurrentRange
            Else
                val = Abs(ids_range_ary(i).hiLimit)  'if pins without bincut spec IDSmax, use current range
            End If
        End If
        
        Call SetCurrentRange(LCase(Pin_Ary(i)), val, WaitTime, ids_range_ary(i).Init_step)
            
    Next i
''**** End - Set init IRange ****
    
    A_HexVS = Split(p_hexvs, ",")
    A_UVS = Split(p_uvs, ",")
    A_VSM = Split(p_vsm, ",")
    
    '''-----------------UFP-----------------
    A_UFP = Split(p_ufp, ",")
    '''-----------------UFP-----------------
    
    'TheHdw.Wait 0.01 'add 10ms.    '20230704 code review no need
    thehdw.Wait WaitTime
    
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Or Profile_byflow Then
        
    Else
        If p_hexvs <> "" Then
            If thehdw.Alarms.Check Then
                HexVS_Power_data = thehdw.DCVS.Pins(p_hexvs).Meter.Read(tlStrobe, 2000, , tlDCVSMeterReadingFormatAverage)
            Else
                HexVS_Power_data = thehdw.DCVS.Pins(p_hexvs).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
            End If
        End If
        If p_uvs <> "" Then UVS_Power_data = thehdw.DCVS.Pins(p_uvs).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
        If p_vsm <> "" Then VSM_Power_data = thehdw.DCVS.Pins(p_vsm).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        If p_ufp <> "" Then UFP_Power_data = thehdw.DCVS.Pins(p_ufp).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
    
    End If
        
    '-------------------------------------------debug print
    ''                debug_print_pins = AutoRange_Pin
    If debug_print_pins <> "" Then
        For i = 0 To UBound(Pin_Ary)
            If InStr(LCase(debug_print_pins), LCase(ids_range_ary(i).PinName)) > 0 Then
                s_InstrumentName = UCase(ids_range_ary(i).ChanMapType)
                For Each site In TheExec.sites
                    Select Case s_InstrumentName
                        Case glbConstIns_HEXVS: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0" & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & HexVS_Power_data.Pins(ids_range_ary(i).PinName).value(site)
                        Case glbConstIns_VHDVS: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0" & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UVS_Power_data.Pins(ids_range_ary(i).PinName).value(site)
                        Case glbConstIns_VSM: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0" & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & VSM_Power_data.Pins(ids_range_ary(i).PinName).value(site)
                        Case glbConstIns_VS5A: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0 " & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UFP_Power_data.Pins(ids_range_ary(i).PinName).value(site) & " Wait time " & WaitTime * 1000
                        Case glbConstIns_VS800MA: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0 " & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UFP_Power_data.Pins(ids_range_ary(i).PinName).value(site) & " Wait time " & WaitTime * 1000
                    End Select
                Next site
            End If
        Next i
    End If
    '-------------------------------------------
                    
    If Search_Step = "" Then
        Stop_Step = 7
    Else
        Stop_Step = CLng(Search_Step)
    End If
    
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
    
    Else
        ''****Start -  Set Auto IRange ****
        If AutoRange_Pin <> "" Then
    
            TheExec.DataManager.DecomposePinList AutoRange_Pin, AutoRangePin_Ary, AutoRangePin_Cnt
    
            For i = 0 To UBound(AutoRangePin_Ary)
                If dic_MeasurePowerPinIndex.Exists(LCase(AutoRangePin_Ary(i))) Then
                    n_Index = dic_MeasurePowerPinIndex(LCase(AutoRangePin_Ary(i)))
                        
                    s_InstrumentName = UCase(ids_range_ary(n_Index).ChanMapType)
                    Select Case s_InstrumentName
                        Case glbConstIns_HEXVS: p_hexvs_autorRange = CombineStringList(p_hexvs_autorRange, AutoRangePin_Ary(i))
                        Case glbConstIns_VHDVS: p_uvs256_autorRange = CombineStringList(p_uvs256_autorRange, AutoRangePin_Ary(i))
                        Case glbConstIns_VSM: p_vsm_autorRange = CombineStringList(p_vsm_autorRange, AutoRangePin_Ary(i))
                        Case glbConstIns_VS5A: p_ufp_autorRange = CombineStringList(p_ufp_autorRange, AutoRangePin_Ary(i))
                        Case glbConstIns_VS800MA: p_ufp_autorRange = CombineStringList(p_ufp_autorRange, AutoRangePin_Ary(i))
                    End Select
                End If
            Next i
    
            If p_vsm_autorRange <> "" Then VSM_Power_data_AutoRange = VSM_Power_data
            If p_uvs256_autorRange <> "" Then UVS256_Power_data_AutoRange = UVS_Power_data
            If p_hexvs_autorRange <> "" Then HexVS_Power_data_AutoRange = HexVS_Power_data
            
            '''-----------------UFP-----------------
            If p_ufp_autorRange <> "" Then UFP_Power_data_AutorRange = UFP_Power_data
            '''-----------------UFP-----------------
            

            For j = 1 To Stop_Step
    
                WaitTime = 260 * us
    
                For i = 0 To UBound(AutoRangePin_Ary)
    
                    Flag_PerPin_autorange = False
                    
                    If dic_MeasurePowerPinIndex.Exists(LCase(AutoRangePin_Ary(i))) Then         '20230807 Tank
                        n_Index = dic_MeasurePowerPinIndex(LCase(AutoRangePin_Ary(i)))
                        
                        s_InstrumentName = UCase(ids_range_ary(n_Index).ChanMapType)
                        Select Case s_InstrumentName
                            Case glbConstIns_HEXVS: PinVal = HexVS_Power_data_AutoRange.Pins(AutoRangePin_Ary(i)).Abs
                            Case glbConstIns_VHDVS: PinVal = UVS256_Power_data_AutoRange.Pins(AutoRangePin_Ary(i)).Abs
                            Case glbConstIns_VSM: PinVal = VSM_Power_data_AutoRange.Pins(AutoRangePin_Ary(i)).Abs
                            Case glbConstIns_VS5A: PinVal = UFP_Power_data_AutorRange.Pins(AutoRangePin_Ary(i)).Abs
                            Case glbConstIns_VS800MA: PinVal = UFP_Power_data_AutorRange.Pins(AutoRangePin_Ary(i)).Abs
                        End Select
                        
                        StepNo = ids_range_ary(n_Index).Init_step - j
                        If StepNo >= 0 Then
                            If ids_range_ary(n_Index).Range_List(StepNo) >= 0.00002 Then
                                DropRngSite = PinVal.compare(LessThan, ids_range_ary(n_Index).Range_List(StepNo) - ids_range_ary(n_Index).Accuracy_List(StepNo + 1))
                                If DropRngSite.Any(True) Then
                                    Flag_PerPin_autorange = True
                                    TheExec.sites.Selected = DropRngSite
                                    thehdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges ids_range_ary(n_Index).Range_List(StepNo), ids_range_ary(n_Index).Range_List(StepNo)
                                    TheExec.sites.Selected = True
                                    SattleTime = ids_range_ary(n_Index).WaitTime_List(StepNo)
                                    If SattleTime > WaitTime Then WaitTime = SattleTime
                                End If
                            End If
                        End If
                        
                    End If
    
                    Set PinVal = Nothing
    
                    Flag_AllPin_autorange = Flag_AllPin_autorange Or Flag_PerPin_autorange
                Next i
    
                If Flag_AllPin_autorange Then
                    
                    thehdw.Wait WaitTime
                    ''Add measurement points to prevent error 20171011 (M9)
                    If p_hexvs_autorRange <> "" Then
                        If thehdw.Alarms.Check Then
                            HexVS_Power_data_AutoRange = thehdw.DCVS.Pins(p_hexvs_autorRange).Meter.Read(tlStrobe, 2000, , tlDCVSMeterReadingFormatAverage)
                        Else
                            HexVS_Power_data_AutoRange = thehdw.DCVS.Pins(p_hexvs_autorRange).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                        End If
                    End If
                    If p_uvs256_autorRange <> "" Then UVS256_Power_data_AutoRange = thehdw.DCVS.Pins(p_uvs256_autorRange).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
                    If p_vsm_autorRange <> "" Then VSM_Power_data_AutoRange = thehdw.DCVS.Pins(p_vsm_autorRange).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                    
                    '''-----------------UFP-----------------
                    If p_ufp_autorRange <> "" Then UFP_Power_data_AutorRange = thehdw.DCVS.Pins(p_ufp_autorRange).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                    '''-----------------UFP-----------------
                    
                    '-------------------------------------------debug print
    ''                debug_print_pins = AutoRange_Pin
                    If debug_print_pins <> "" Then
                        For i = 0 To UBound(AutoRangePin_Ary)
                            If dic_MeasurePowerPinIndex.Exists(LCase(AutoRangePin_Ary(i))) Then     'If InStr(LCase(debug_print_pins), LCase(ids_range_ary(i).PinName)) > 0 Then
                                s_InstrumentName = UCase(ids_range_ary(i).ChanMapType)
                                For Each site In TheExec.sites
                                    Select Case s_InstrumentName
                                        Case glbConstIns_HEXVS: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).Meter.CurrentRange.value & ", Current: " & HexVS_Power_data_AutoRange.Pins(ids_range_ary(i).PinName).value(site)
                                        Case glbConstIns_VHDVS: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UVS256_Power_data_AutoRange.Pins(ids_range_ary(i).PinName).value(site)
                                        Case glbConstIns_VSM: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & VSM_Power_data_AutoRange.Pins(ids_range_ary(i).PinName).value(site)
                                        Case glbConstIns_VS5A: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UFP_Power_data_AutorRange.Pins(ids_range_ary(i).PinName).value(site)
                                        Case glbConstIns_VS800MA: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & thehdw.DCVS.Pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UFP_Power_data_AutorRange.Pins(ids_range_ary(i).PinName).value(site)
                                    End Select
                                Next site
                            End If
                        Next i
                    End If
                    '-------------------------------------------
                    Flag_AllPin_autorange = False
                Else
    
                    j = Stop_Step
    
                End If
    
            Next j
    
        End If
        
    ''****Start -  Set Auto IRange ****
    End If

    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Or Profile_byflow Then
        For i = 0 To UBound(Pin_Ary)
            Power_data.AddPin Pin_Ary(i)
        Next i
        
    Else
        For i = 0 To UBound(A_HexVS)
            Power_data.AddPin (A_HexVS(i))
            If LCase("*," & p_hexvs_autorRange & ",*") Like LCase("*," & A_HexVS(i) & ",*") Then
                Power_data.Pins(A_HexVS(i)) = HexVS_Power_data_AutoRange.Pins(A_HexVS(i))
            Else
                Power_data.Pins(A_HexVS(i)) = HexVS_Power_data.Pins(A_HexVS(i))
            End If
            'offline mode simulation
''''            If TheExec.TesterMode = testModeOffline Then
''''                For Each Site In TheExec.sites
''''                    Power_data.Pins(A_HexVS(i)).value(Site) = 0.01 + Rnd() * 0.0001
''''                Next Site
''''            End If
        Next i
    
        For i = 0 To UBound(A_UVS)
            Power_data.AddPin (A_UVS(i))
            If LCase("*," & p_uvs256_autorRange & ",*") Like LCase("*," & A_UVS(i) & ",*") Then
                Power_data.Pins(A_UVS(i)) = UVS256_Power_data_AutoRange.Pins(A_UVS(i))
            Else
                Power_data.Pins(A_UVS(i)) = UVS_Power_data.Pins(A_UVS(i))
            End If
            'offline mode simulation
''''            If TheExec.TesterMode = testModeOffline Then
''''                For Each Site In TheExec.sites
''''                    Power_data.Pins(A_UVS(i)).value(Site) = 0.0005 + Rnd() * 0.0001
''''                Next Site
''''            End If
        Next i
    
        For i = 0 To UBound(A_VSM)
            Power_data.AddPin (A_VSM(i))
            If LCase("*," & p_vsm_autorRange & ",*") Like LCase("*," & A_VSM(i) & ",*") Then
                Power_data.Pins(A_VSM(i)) = VSM_Power_data_AutoRange.Pins(A_VSM(i))
            Else
                Power_data.Pins(A_VSM(i)) = VSM_Power_data.Pins(A_VSM(i))
            End If
            'offline mode simulation
''''            If TheExec.TesterMode = testModeOffline Then
''''                For Each Site In TheExec.sites
''''                    Power_data.Pins(A_VSM(i)).value(Site) = 0.0005 + Rnd() * 0.0001
''''                Next Site
''''            End If
        Next i
    
    '''-----------------UFP-----------------
        For i = 0 To UBound(A_UFP)
            Power_data.AddPin (A_UFP(i))
            If LCase("*," & p_ufp_autorRange & ",*") Like LCase("*," & A_UFP(i) & ",*") Then
                Power_data.Pins(A_UFP(i)) = UFP_Power_data_AutorRange.Pins(A_UFP(i))
            Else
                Power_data.Pins(A_UFP(i)) = UFP_Power_data.Pins(A_UFP(i))
            End If
            
            'offline mode simulation
''''            If TheExec.TesterMode = testModeOffline Then
''''                For Each Site In TheExec.sites
''''                    Power_data.Pins(A_UFP(i)).value(Site) = 0.0005 + Rnd() * 0.0001
''''                Next Site
''''            End If
            Next i
            '''-----------------UFP-----------------
    End If

    If isUse_Product_Identifier Then
                '//////////////////////////////////////////if is seudofuse//////////////////////////////////////////
                    
        If PseudoFuseEnable Then
            If BdfDataBase.Banks("CFG").Fields.Exists("product_identifier") Then
                Set field_normal = BdfDataBase.Banks("CFG").Fields("product_identifier")
                For Each site In TheExec.sites
                    CurrentPassBinCutNum_normal(site) = field_normal.DsscDecValue + 1
                    If CurrentPassBinCutNum_normal(site) > Total_Bincut_Num Then
                        GlbUtility.WriteDlg "site:" & site & ", product_identifier " & CurrentPassBinCutNum_normal(site) & " > Total_Bincut_Num " & Total_Bincut_Num & " , Error!!!"
                        TheExec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="product_identifier Error"
                    End If
                Next site
            Else
                CurrentPassBinCutNum_normal = 1
            End If
        End If
    '//////////////////////////////////////////if is seudofuse//////////////////////////////////////////

        nBinCutNumber = CurrentPassBinCutNum_normal
    Else
        nBinCutNumber = CLng(nSpecific_Product_Identifier)
    End If
    
    For Each site In TheExec.sites
        nTempBinCutNum(site) = CheckHaveBinCutSheet(nBinCutNumber(site), isUse_Product_Identifier)
    Next site
    
    For i = 0 To CorePowerPin_Cnt - 1: For j = 0 To repeat_count - 1
        If gl_GetInstrument_Dic.Exists(LCase(CorePowerPin_Ary(i))) Then
            
            Tname = Report_TName_From_Instance("I", CorePowerPin_Ary(i), , , , , , , tlForceNone)
            VMain = Format(thehdw.DCVS.Pins(Power_data.Pins(CorePowerPin_Ary(i))).Voltage.Main.value, "0.00")
            Hilimit_IDS = 0
            Lolimit_IDS = 0
            If dic_MeasurePowerPinIndex.Exists(LCase(CorePowerPin_Ary(i))) Then
                n_Index = dic_MeasurePowerPinIndex(LCase(CorePowerPin_Ary(i)))

                For Each site In TheExec.sites
                    Hilimit_IDS = Compare_ForceVal_BV(Pin_Ary(n_Index), Val_Hi(n_Index), True, isUse_Product_Identifier, nTempBinCutNum(site))
                    Lolimit_IDS = Compare_ForceVal_BV(Pin_Ary(n_Index), Val_Lo(n_Index), True, isUse_Product_Identifier, nTempBinCutNum(site))
                    If TheExec.TesterMode = testModeOffline Or gl_EnableCurrentProfile Or gl_EnableVoltageProfile Or Profile_byflow Then
                        d_SimulationValue = (Hilimit_IDS + Lolimit_IDS) / 2
                        Power_data.Pins(Pin_Ary(n_Index)).value(site) = d_SimulationValue
                    End If
                    
                    If ENG_Limit = False Then
                        TheExec.Datalog.WriteComment "Site(" & site & ") power pin : " & CorePowerPin_Ary(i) & " - value : " & Format((Power_data.Pins(CorePowerPin_Ary(i)).value(site) * 1000), "0.000000") & " mA"
                    Else
                    
                        If Hilimit_IDS = 0 Then     'If Hilimit_IDS = 0 And Lolimit_IDS = 0 Then
                            TheExec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, highCompareSign:=tlSignNone       ', highCompareSign:=tlSignNone, lowCompareSign:=tlSignNone
                        Else
                            TheExec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, hiVal:=Hilimit_IDS   'theexec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, hiVal:=Hilimit_IDS, lowVal:=Lolimit_IDS
                        End If
                        
                    End If
                    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
                Next site

            End If
            
            TheExec.Datalog.WriteComment "Current I Range: " & CorePowerPin_Ary(i) & "--->" & thehdw.DCVS.Pins(CorePowerPin_Ary(i)).Meter.CurrentRange.value
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1

        End If
    Next j: Next i

    For k = 0 To OtherPowerPin_Cnt - 1
        If gl_GetInstrument_Dic.Exists(LCase(OtherPowerPin_Ary(k))) Then

            'Tname = theexec.DataManager.instanceName & "_" & OtherPowerPin_Ary(k)        'add pin name Aruba 2017/12/28
            Tname = Report_TName_From_Instance("I", OtherPowerPin_Ary(k), , , , , , , tlForceNone)
            VMain = Format(thehdw.DCVS.Pins(Power_data.Pins(OtherPowerPin_Ary(k))).Voltage.Main.value, "0.00")

            Hilimit_IDS = 0
            Lolimit_IDS = 0
            
            If dic_MeasurePowerPinIndex.Exists(LCase(OtherPowerPin_Ary(k))) Then
                n_Index = dic_MeasurePowerPinIndex(LCase(OtherPowerPin_Ary(k)))

                For Each site In TheExec.sites
                    Hilimit_IDS = Compare_ForceVal_BV(Pin_Ary(n_Index), Val_Hi(n_Index), True, isUse_Product_Identifier, nTempBinCutNum(site))
                    Lolimit_IDS = Compare_ForceVal_BV(Pin_Ary(n_Index), Val_Lo(n_Index), True, isUse_Product_Identifier, nTempBinCutNum(site))
                    If TheExec.TesterMode = testModeOffline Or gl_EnableCurrentProfile Or gl_EnableVoltageProfile Or Profile_byflow Then
                        d_SimulationValue = (Hilimit_IDS + Lolimit_IDS) / 2
                        Power_data.Pins(Pin_Ary(n_Index)).value(site) = d_SimulationValue
                    End If
                    
                    If ENG_Limit = False Then
                        TheExec.Datalog.WriteComment "Site(" & site & ") power pin : " & OtherPowerPin_Ary(k) & " - value : " & Format((Power_data.Pins(OtherPowerPin_Ary(k)).value(site) * 1000), "0.000000") & " mA"
                    Else
                    
                        If Hilimit_IDS = 0 Then     'If Hilimit_IDS = 0 And Lolimit_IDS = 0 Then
                            TheExec.Flow.TestLimit resultVal:=Power_data.Pins(OtherPowerPin_Ary(k)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, highCompareSign:=tlSignNone
                        Else
                            TheExec.Flow.TestLimit resultVal:=Power_data.Pins(OtherPowerPin_Ary(k)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, hiVal:=Hilimit_IDS
                        End If
                        
                    End If
                    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
                Next site

            End If
            
            TheExec.Datalog.WriteComment "Current I Range: " & OtherPowerPin_Ary(k) & "--->" & thehdw.DCVS.Pins(OtherPowerPin_Ary(k)).Meter.CurrentRange.value
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1

        End If
    Next k

    'recover range setup
    
    'Pin_Cnt = CorePowerPin_Cnt + OtherPowerPin_Cnt
    
    For i = 0 To UBound(Pin_Ary)
        thehdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges ids_range_ary(i).Init_CurrentRange, ids_range_ary(i).Init_CurrentRange
    Next i
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_IDS", "DCVS_IDS_main_auto_range_and_measure") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function DCVI_IDS_main_auto_range_and_measure(CorePower_Pin As String, _
                                                         Power_data As PinListData, _
                                                         repeat_count As Long, _
                                                         FlowLimitForInitIRange As Boolean, _
                                                         Optional Search_Step As String, _
                                                         Optional debug_print_pins As String, _
                                                         Optional AutoRange_Pin As String, _
                                                         Optional FixCurrentRange As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim i As Long, j As Long, All_Power_Pin As String
    Dim site As Variant, pin As Variant, val As Double
    Dim k As Long
    Dim Tname As String
    Dim VMain As Double
    Dim p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim CorePowerPin_Ary() As String, CorePowerPin_Cnt As Long
    Dim OtherPowerPin_Ary() As String, OtherPowerPin_Cnt As Long
    
    Dim ChannelType As Long
    Dim Channels() As String, NumberChannels As Long
    Dim NumberSites As Long, Error As String
    
    Dim Powerpin_log As String ''20180315 Abel added
                                                                                                                                                                                                                                           
    Dim funcName As String:: funcName = "IDS_main_auto_range_and_measure"
    
    'Get the limits info
    Dim FlowLimitsInfo As IFlowLimitsInfo
    
    All_Power_Pin = vbNullString
    Tname = vbNullString
    Error = vbNullString
    
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    
    'if no Use-Limits on this test, FlowLimitsInfo is nothing
    If FlowLimitsInfo Is Nothing Then
        If isDebugMode Then TheExec.AddOutput "Could not get the limits info", vbRed, True
        Exit Function
    Else
    End If

    Dim Val_Hi() As String
    Dim Val_Lo() As String
    FlowLimitsInfo.GetHighLimits Val_Hi
    FlowLimitsInfo.GetLowLimits Val_Lo
    
    Dim dcvs_pincnt As Long
    dcvs_pincnt = Power_data.Pins.Count

    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim Merge_Type, Slot_Type As String
    Dim Split_Ary() As String
    Dim SattleTime As Double
    Dim WaitTime As Double
    Dim p_hexvs As String
    Dim p_uvs As String
    Dim A_HexVS() As String
    Dim A_UVS() As String
    Dim HexVS_Power_data As New PinListData
    Dim UVS_Power_data As New PinListData

    ''For UVI80 Irange:1A/200mA/20mA/2mA/200uA (2017/7/31)
    Dim p_uvi80 As String
    Dim A_UVI80() As String
    Dim UVI80_Power_data As New PinListData

    Dim IDS_ini_Current_range() As Double

    Dim SlotType As Scripting.Dictionary
    Set SlotType = New Scripting.Dictionary
    Dim InitStep As Scripting.Dictionary
    Set InitStep = New Scripting.Dictionary
    Dim PinVal As New PinData
    Dim DropRngSite As New SiteBoolean
    Dim AutoRangePin_Ary() As String
    Dim AutoRangePin_Cnt As Long
    Dim range_ary() As AutoRange_Info
    Dim dict_FixedCurrent As New Scripting.Dictionary
    Dim p_uvi80_autorRange As String
    Dim A_uvi80_autorRange() As String
    Dim UVI80_Power_data_AutoRange As New PinListData
    Dim Flag_PerPin_autorange As Boolean
    Dim Flag_AllPin_autorange As Boolean
    
    p_uvi80 = vbNullString
    Merge_Type = vbNullString
    Slot_Type = vbNullString
    p_uvi80_autorRange = vbNullString
    
    TheExec.DataManager.DecomposePinList CorePower_Pin, CorePowerPin_Ary, CorePowerPin_Cnt

    For i = 0 To CorePowerPin_Cnt - 1
        If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) <> "N/C" Then
            All_Power_Pin = All_Power_Pin & "," & CorePowerPin_Ary(i)
        Else
        End If
    Next i

    If All_Power_Pin <> "" Then
        All_Power_Pin = right(All_Power_Pin, Len(All_Power_Pin) - 1)
    Else
    End If
    
    Pin_Ary = Split(All_Power_Pin, ",")
    ReDim IDS_ini_Current_range(UBound(Pin_Ary)) As Double
    WaitTime = 100 * us
    
    'Add for Fixcurrent feature michael 20240522
    Call FixCurrentRange_StrToDic(FixCurrentRange, dict_FixedCurrent)
    
    ''**** Start - Get PowerPin Info ****
    Dim ids_range_ary() As AutoRange_Info
    ReDim ids_range_ary(UBound(Pin_Ary))
    For i = 0 To UBound(Pin_Ary)
        For j = 0 To UBound(PowerPin_range_ary)
            If Pin_Ary(i) = PowerPin_range_ary(j).PinName Then
                ids_range_ary(i) = PowerPin_range_ary(j)
            Else
            End If
        Next j
    Next i
    ''**** End - Get PowerPin Info ****

    ' Set init IRange
    For i = 0 To UBound(Pin_Ary)
    
        ids_range_ary(i).Init_CurrentRange = thehdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.value
        ids_range_ary(i).Init_Source_FoldLimit = thehdw.DCVI.Pins(Pin_Ary(i)).Current
        ids_range_ary(i).hiLimit = Compare_ForceVal_BV(Pin_Ary(i), Val_Hi(i + dcvs_pincnt)) 'if pins without limit use bincut spec IDSmax as current range
    
        If LCase(ids_range_ary(i).ChanMapType) = "dc-07" Then
            p_uvi80 = p_uvi80 & "," & Pin_Ary(i)
            thehdw.DCVI.Pins(Pin_Ary(i)).Meter.mode = tlDCVIMeterCurrent '20180115 Rick
        Else
        End If
        
        
        If dict_FixedCurrent.Exists(LCase(Pin_Ary(i))) Then
            val = CDbl(dict_FixedCurrent(LCase(Pin_Ary(i))))
        ElseIf (FlowLimitForInitIRange = False) Then
            val = thehdw.DCVI.Pins(Pin_Ary(i)).Current 'InitIRange_Setup ids_range_ary(i), WaitTime, SattleTime
        Else
            If ids_range_ary(i).hiLimit = 0 Then
                val = ids_range_ary(i).Init_CurrentRange
            Else
                val = Abs(ids_range_ary(i).hiLimit) 'if pins without bincut spec IDSmax, use current range
            End If
            'If Val_Hi(i) = "" Then Val = range_ary(i).Init_CurrentRange Else Val = Abs(Val_Hi(i)) 'if pins without limit use current range
            'If TheExec.Flow.EnableWord("Temp_25C") Or TheExec.Flow.EnableWord("Temp_N25C") Then
            If val < 0.0002 Then
                val = 0.0002
            Else
            End If
            

        End If
        Call SetCurrentRange(LCase(Pin_Ary(i)), val, WaitTime, ids_range_ary(i).Init_step)
'        Debug.Print i & ":" & CorePower_Pin_Ary(i)
    Next i

    If p_uvi80 <> "" Then p_uvi80 = right(p_uvi80, Len(p_uvi80) - 1)
    A_UVI80 = Split(p_uvi80, ",")

    thehdw.Wait 0.035 'add 10ms.
    thehdw.Wait WaitTime

    ReDim AutoRangePin(UBound(CorePowerPin_Ary))
    AutoRangePin_Ary = CorePowerPin_Ary

    If p_uvi80 <> "" Then
        UVI80_Power_data = thehdw.DCVI.Pins(p_uvi80).Meter.Read(tlStrobe, 1, , tlDCVIMeterReadingFormatAverage)
    Else
    End If
    '-------------------------------------------debug print, disable with HIP TTR enable word
''    debug_print_pins = AutoRange_Pin
    If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False And gl_Disable_IDS_AutoRange_log = False Then
        If debug_print_pins <> "" Then
            For i = 0 To UBound(Pin_Ary)
                For k = 0 To UBound(ids_range_ary)
                    If UCase(Pin_Ary(i)) = UCase(ids_range_ary(k).PinName) Then
                        If InStr(LCase(debug_print_pins), LCase(ids_range_ary(k).PinName)) > 0 Then
                            For Each site In TheExec.sites
                                If LCase(ids_range_ary(k).ChanMapType) = "dc-07" Then
                                    TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(k).PinName & ", Step 0" & ", Irange: " & thehdw.DCVI.Pins(ids_range_ary(k).PinName).Meter.CurrentRange.value & ", Current: " & UVI80_Power_data.Pins(ids_range_ary(k).PinName).value(site)
                                Else
                                End If
                            Next site
                        Else
                        End If
                    Else
                    End If
                Next k
            Next i
        Else
        End If
    Else
    End If
    '-------------------------------------------
    
    Dim Stop_Step, StepNo As Integer
    If Search_Step = "" Then
        Stop_Step = 5
    ElseIf (CLng(Search_Step) >= 6) Then
        Stop_Step = 6
    Else
        Stop_Step = CLng(Search_Step)
    End If
    '========================================================================================auto range search

    
    If TheExec.Flow.enableWord("CurrentProfile") Or TheExec.Flow.enableWord("VoltageProfile") Or Profile_byflow = True Then
    
    Else
        ''****Start -  Set Auto IRange ****
        If AutoRange_Pin <> "" Then
        
            TheExec.DataManager.DecomposePinList AutoRange_Pin, AutoRangePin_Ary, AutoRangePin_Cnt
    
            For i = 0 To UBound(AutoRangePin_Ary)
                For k = 0 To UBound(ids_range_ary)
                    If LCase(AutoRangePin_Ary(i)) = LCase(ids_range_ary(k).PinName) Then
                        If LCase(ids_range_ary(k).ChanMapType) = "dc-07" Then
                            p_uvi80_autorRange = p_uvi80_autorRange & "," & AutoRangePin_Ary(i)
                        Else
                        End If
                        Exit For
                    Else
                    End If
                Next k
            Next i
    
            If p_uvi80_autorRange <> "" Then
                p_uvi80_autorRange = right(p_uvi80_autorRange, Len(p_uvi80_autorRange) - 1)
                A_uvi80_autorRange = Split(p_uvi80_autorRange, ",")
                UVI80_Power_data_AutoRange = UVI80_Power_data
            Else
            End If

            For j = 1 To Stop_Step
    
                WaitTime = 260 * us
    
                For i = 0 To UBound(AutoRangePin_Ary)
    
                    Flag_PerPin_autorange = False
    
                    For k = 0 To UBound(ids_range_ary)
    
                        If LCase(AutoRangePin_Ary(i)) = LCase(ids_range_ary(k).PinName) Then
    
                            If LCase(ids_range_ary(k).ChanMapType) = "dc-07" Then
                                PinVal = UVI80_Power_data_AutoRange.Pins(AutoRangePin_Ary(i)).Abs
                            Else
                            End If
                            
                            StepNo = ids_range_ary(k).Init_step - j
                            If StepNo >= 0 Then
                                If ids_range_ary(k).Range_List(StepNo) >= 0.00002 Then ''20230710, If current range is equal and larger than 20uA, AutoRange will be enabled.
                                    DropRngSite = PinVal.compare(LessThan, ids_range_ary(k).Range_List(StepNo) - ids_range_ary(k).Accuracy_List(StepNo))
                                    'DropRngSite = PinVal.Compare(LessThan, ids_range_ary(k).Range_List(StepNo) * 0.5) 'KC fix alarm issue 211217
                                    If DropRngSite.Any(True) Then
                                        Flag_PerPin_autorange = True
                                        TheExec.sites.Selected = DropRngSite
                                        thehdw.DCVI.Pins(AutoRangePin_Ary(i)).SetCurrentAndRange ids_range_ary(k).Range_List(StepNo), ids_range_ary(k).Range_List(StepNo)
                                        TheExec.sites.Selected = True
                                        SattleTime = ids_range_ary(k).WaitTime_List(StepNo)
                                        If SattleTime > WaitTime Then
                                            WaitTime = SattleTime
                                        Else
                                        End If
                                    Else
                                    End If
                                Else
                                End If
                            Else
                            End If
                            
                            Exit For
                        Else
                        End If
    
                    Next k
    
                    Set PinVal = Nothing
    
                    Flag_AllPin_autorange = Flag_AllPin_autorange Or Flag_PerPin_autorange
                Next i
    
                If Flag_AllPin_autorange Then
                    'If TheExec.Flow.EnableWord("Temp_85C") Then WaitTime = WaitTime + 0.015
                    thehdw.Wait WaitTime

                    'enable mode alarm for current range change on DCVI power pin. 20211224
                    thehdw.DCVI.Pins(p_uvi80_autorRange).Alarm(tlDCVIAlarmMode) = tlAlarmForceFail

                    If p_uvi80_autorRange <> "" Then
                        UVI80_Power_data_AutoRange = thehdw.DCVI.Pins(p_uvi80_autorRange).Meter.Read(tlStrobe, 32, , tlDCVIMeterReadingFormatAverage)
                    Else
                    End If
                    '-------------------------------------------debug print, disable with HIP TTR enable word
                    If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False And gl_Disable_IDS_AutoRange_log = False Then
    ''                debug_print_pins = AutoRange_Pin
                        If debug_print_pins <> "" Then
                            For i = 0 To UBound(AutoRangePin_Ary)
                                For k = 0 To UBound(ids_range_ary)
                                    If UCase(AutoRangePin_Ary(i)) = UCase(ids_range_ary(k).PinName) Then
                                        If InStr(LCase(debug_print_pins), LCase(ids_range_ary(k).PinName)) > 0 Then
                                            For Each site In TheExec.sites
                                                If LCase(ids_range_ary(k).ChanMapType) = "dc-07" Then
                                                    TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(k).PinName & ", Step " & j & ", Irange: " & thehdw.DCVI.Pins(ids_range_ary(k).PinName).CurrentRange.value & ", Current: " & UVI80_Power_data_AutoRange.Pins(ids_range_ary(k).PinName).value(site)
                                                Else
                                                End If
                                            Next site
                                        Else
                                        End If
                                    Else
                                    End If
                                Next k
                            Next i
                        Else
                        End If
                    Else
                    End If
                    '-------------------------------------------
                    Flag_AllPin_autorange = False
                Else
    
                    j = Stop_Step
    
                End If
    
            Next j
        Else
        End If
        
    ''****Start -  Set Auto IRange ****
    End If
    '========================================================================================auto range search

    For i = 0 To UBound(A_UVI80)
        Power_data.AddPin (A_UVI80(i))
        If LCase("*," & p_uvi80_autorRange & ",*") Like LCase("*," & A_UVI80(i) & ",*") Then
            Power_data.Pins(A_UVI80(i)) = UVI80_Power_data_AutoRange.Pins(A_UVI80(i))
        Else
            Power_data.Pins(A_UVI80(i)) = UVI80_Power_data.Pins(A_UVI80(i))
        End If
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                Power_data.Pins(A_UVI80(i)).value(site) = 0.01 + Rnd() * 0.0001
            Next site
        Else
        End If
    Next i

    For i = 0 To CorePowerPin_Cnt - 1: For j = 0 To repeat_count - 1
        If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) <> "N/C" Then

            ''20180315 Abel change naming'            Tname = TheExec.DataManager.InstanceName & "_" & j
'            Tname = TheExec.DataManager.instanceName 'No need _0
            Tname = Report_TName_From_Instance("I", Power_data.Pins(CorePowerPin_Ary(i)), , , , , , , tlForceNone)
            If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) Like "*DCVI*" Then
                VMain = Format(thehdw.DCVI.Pins(Power_data.Pins(CorePowerPin_Ary(i))).Voltage, "0.00")
            Else
            End If
            TheExec.Datalog.WriteComment (glb_TestInstance & " =====> Curr_meas Meter I range setting, " & CorePowerPin_Ary(i) & " =" & thehdw.DCVI.Pins(CorePowerPin_Ary(i)).Meter.CurrentRange.value)

            If TPModeAsCharz_GLB = True Then 'wc 180319 for charz
                TheExec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            Else
                Powerpin_log = Replace(UCase(CorePowerPin_Ary(i)), "_", "")
                TheExec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow   ''20180315 Abel add power pin name
            End If
             TheExec.Datalog.WriteComment "Current I Range: " & CorePowerPin_Ary(i) & "--->" & thehdw.DCVI.Pins(CorePowerPin_Ary(i)).Meter.CurrentRange.value
        Else
        End If
    Next j: Next i

    For i = 0 To UBound(Pin_Ary)
        If TheExec.DataManager.ChannelType(Pin_Ary(i)) Like "*DCVI*" Then
            thehdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange ids_range_ary(i).Init_Source_FoldLimit, ids_range_ary(i).Init_CurrentRange
        Else
        End If
    Next i

    thehdw.Wait 0.003

    Exit Function
    
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_IDS", "DCVI_IDS_main_auto_range_and_measure") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function





' [20231003][All][Tank] modify after Chihome review
' [20231016][All][Tank] Use ApplyLevelsTiming to set iFold limit
Public Function IDS_main_current(patt As Pattern, _
                                      DCVS_Power_Pin As String, DCVI_Power_Pin As String, _
                                      DCVS_OtherPower_Pin As String, _
                                      repeat_count As Long, _
                                      FlowLimitForInitIRange As Boolean, _
                             Optional Search_Step As String, _
                             Optional DisableClock As Boolean = False, _
                             Optional FlagWait As Boolean = False, _
                             Optional FixCurrentRange As String, _
                             Optional CharInputString As String, Optional BV_SpecialHandle As String, Optional DisableClockPortName As String, Optional DisableFRCPinName As String, Optional FRC_RelayPin As String, _
                             Optional RTOSCmd As String = vbNullString, Optional RTOSTimeOut As Double, Optional DisconnectClock As Boolean = True, Optional debug_print_pins As String, Optional NotPrintOutLimit As Boolean = False, Optional Interpose_Meas_before As String, Optional Interpose_Meas_after As String, _
                             Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, Optional CUS_Str_DigSrcData As String = vbNullString, _
                             Optional DictName As String, Optional Fuse_Enable As Boolean, Optional Calc_Eqn As String, Optional AutoRange_Pin As String, Optional Validating_ As Boolean, _
                             Optional isUse_Product_Identifier As Boolean = False, Optional sSpecific_Product_Identifier As String = "999", Optional EnableDisconnectPins As Boolean = False, Optional ExcludeDisconnectPins As String = vbNullString)    'Carter, 20190315     'FixCurrentRange 20220627       '20221014 Tank add bincut sheet reference product identifier

On Error GoTo errHandler
    'gl_IDS_INFO_Dic.RemoveAll
    Dim p As Variant
    Dim p_ary() As String
    Dim PinCnt As Long
    Dim MeasCurr_HexVS As New PinListData
    Dim MeasCurr As New PinListData
    Dim MeasCurr_copy As New PinListData
    Dim Power_pin As String
    Dim testnum() As Long, Cnt1 As Long
    Dim i As Long, j As Long
    Dim repeat_judge As Long

    Dim All_Power_data As New PinListData
    Dim site As Variant

    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim CLK_Pins As String

    Dim rtnPatNames() As String
    Dim PatCnt As Long
    Dim InDSPWave As New DSPWave
    Dim temp_instance_name As String: temp_instance_name = TheExec.DataManager.InstanceName

    Dim sTempCalc_Eqn As String

    Dim sOutPut_DisconnectPin As String
    Dim sOutput_IgnorePin As String
    Dim sDontCarePin As String
    Dim sOutput_NCPin As String
    Dim isDoDisconnectPinPass As Boolean
    
    Dim s_DebugInfo(6) As String
    Dim s_DebugInfoString As String
    
    Dim sFuncName As String:: sFuncName = "IDS_main_current"
    Dim All_Power_Pin As String
    
    glb_TestInstance = UCase(TheExec.DataManager.InstanceName)
    
    Power_pin = vbNullString
    CLK_Pins = vbNullString
    sTempCalc_Eqn = vbNullString
    sOutPut_DisconnectPin = vbNullString
    sDontCarePin = vbNullString
    sOutput_NCPin = vbNullString
    s_DebugInfoString = vbNullString
    All_Power_Pin = vbNullString
    
    If Validating_ Then 'Carter, 20190315
        Call PrLoadPattern(patt.value)
        Exit Function    ' Exit after validation
    Else
    End If

    
    If isUse_Product_Identifier Then    'replace Calc_Eqn string to add production identifier
        sTempCalc_Eqn = Replace(Calc_Eqn, "~", "~d")    'use efuse value in device
    Else
        If IsNumeric(sSpecific_Product_Identifier) = False Then
            sSpecific_Product_Identifier = "999"
        Else
        End If
        sTempCalc_Eqn = Replace(Calc_Eqn, "~", "~" & sSpecific_Product_Identifier)      'use specific production identifier
    End If
    
    If ENG_SweepPin Then
        TheExec.Datalog.WriteComment "[" & temp_instance_name & "  start]"
        TheExec.Datalog.WriteComment "Threshold_CriticalPin =" & Threshold_CriticalPin
        ENG_Limit = False
        Call FindInputPin(patt, FlagWait, DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData)

        Call SweepPinInPattern(patt, FlagWait, DCVS_Power_Pin, DCVS_OtherPower_Pin, repeat_count, FlowLimitForInitIRange, _
            DisableClock, Search_Step, DisableClockPortName, DisconnectClock, debug_print_pins, AutoRange_Pin, _
            DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, _
            Interpose_Meas_before, Interpose_Meas_after, DisableFRCPinName, FRC_RelayPin)

        TheExec.Datalog.WriteComment "[" & temp_instance_name & "  end]"
    Else
    End If
    ENG_Limit = True
    
    
    thehdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag

    thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    'Add SPI init cindition. Merge NAND IDS and SPI IDS in this module. 20160903 ylliuj
'    If RTOS_Setup = True Then
'        SPI_Initial_Conds_Fun
'    End If

'    If UCase(TheExec.CurrentJob) Like "*CHAR*" Then
'        If CharInputString <> "" Then
'            Call SetForceCondition(CharInputString)
'        End If
'    End If

    'TheHdw.Digital.Patgen.TimeoutEnable = False
    
    gb_IDSLimit_Speical_Handle = False
    If BV_SpecialHandle <> "" Then
        gDict_IDSLimit_Special_Handle.RemoveAll
        IDS_ExtraLimit BV_SpecialHandle
    Else
    End If

    sOutPut_DisconnectPin = vbNullString
    sOutput_IgnorePin = vbNullString
    sDontCarePin = vbNullString
    sOutput_NCPin = vbNullString
    
    isDoDisconnectPinPass = Disconnect_X_and_Nouse_Pin(DisconnectPinType.UnusedIoPin, patt, EnableDisconnectPins, ExcludeDisconnectPins, sDontCarePin, sOutPut_DisconnectPin, sOutput_IgnorePin, sOutput_NCPin)     'Disconnect UnUse pin after run pattern
    If isDoDisconnectPinPass = False Then
        Exit Function     'Binout if get something error no need do IDS test
    Else
    End If
    
    thehdw.Digital.Patgen.TimeOut = 10
    
    If patt <> "" Then
        'Call TheHdw.Patterns(patt).Load    '20230406 pattern load is doing when validated, so no need to do again.
        '-------------------------------------------DSSC Soruce
        If DigSrc_pin <> "" Then
            rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(patt, PatCnt)
            Call GeneralDigSrcSetting(CStr(rtnPatNames(0)), DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, _
                                    digsrc_assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPWave)
        Else
        End If
        '-------------------------------------------
    
        If FlagWait Then
            Call thehdw.Patterns(patt).start
            Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)  'Meas during CPUA loop
        Else
            Call thehdw.Patterns(patt).TEST(pfAlways, 0, tlResultModeDomain)
            thehdw.Digital.Patgen.HaltWait
        End If
    
    Else
        '------------------------------ 191017 Tonga BringUp RTOS Run Scenario Start by Leslie---------------------------------------------------------------------------
        If RTOSCmd <> "" Then
            RTOS_IDS RTOSCmd, RTOSTimeOut
        End If
        '------------------------------ 191017 Tonga BringUp RTOS Run Scenario Start by Leslie---------------------------------------------------------------------------
    End If
    '-------------------------------------------AP & RF
    
    'Call SetCriticalPins_InitLo(CriticalPin_InitLo)
    'Call SetCriticalPins_RelayOff(CriticalPin_RelayOff)
    'Call SetCriticalPins_VtMode(CriticalPin_VtMode)
    
    If DisableClock Then
        If DisableFRCPinName <> "" And FRC_RelayPin <> "" Then
            Call Disable_FRC_Pins(DisableFRCPinName, FRC_RelayPin, DisableClock)
        Else
            Call Disable_FRC(DisableClockPortName, DisconnectClock)
        End If
        Wait 0.005
    Else
    End If
        
    Call Disconnect_X_and_Nouse_Pin(DisconnectPinType.Xpin, patt, EnableDisconnectPins, sDontCarePin:=sDontCarePin)   'Disconnect Xpin after run pattern
    
    '-------------------------------------------pre_measure_store
    If Interpose_Meas_before <> "" Then
        Call SetForceCondition(Interpose_Meas_before & ";STOREPREMEAS")
    Else
    End If
    '-------------------------------------------

    Set All_Power_data_IDS_GB = Nothing
    If DCVS_Power_Pin <> "" Or DCVS_OtherPower_Pin <> "" Then
        If (UCase(DCVS_Power_Pin) Like "*CP*=*" Or UCase(DCVS_Power_Pin) Like "*FT*=*") Then
            DCVS_Power_Pin = Select_MeasPin(DCVS_Power_Pin, UCase(currentJobName))
        Else
        End If
        If (UCase(AutoRange_Pin) Like "*CP*=*" Or UCase(AutoRange_Pin) Like "*FT*=*") Then
            AutoRange_Pin = Select_MeasPin(AutoRange_Pin, UCase(currentJobName))
        Else
        End If
        DCVS_IDS_main_auto_range_and_measure DCVS_Power_Pin, DCVS_OtherPower_Pin, All_Power_data, repeat_count, FlowLimitForInitIRange, Search_Step, debug_print_pins, AutoRange_Pin, FixCurrentRange, isUse_Product_Identifier, sSpecific_Product_Identifier        'FixCurrentRange 20220627
    Else
    End If
    If DCVI_Power_Pin <> "" Then
        DCVI_IDS_main_auto_range_and_measure DCVI_Power_Pin, All_Power_data, repeat_count, FlowLimitForInitIRange, Search_Step, debug_print_pins, AutoRange_Pin, FixCurrentRange
    Else
    End If

    All_Power_data_IDS_GB = All_Power_data.Copy
    Wait 0.005

'''    Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
'''    '============================== For efuse =================================
'''    'collect measure values
'''    If eFusePower_Pin <> "" Then
'''        TheExec.DataManager.DecomposePinList eFusePower_Pin, p_ary, PinCnt
'''        For Each Site In TheExec.sites
'''            For i = 0 To PinCnt - 1
'''                'For eFuse category naming rule was fixed
'''                Call auto_eFuse_IDS_SetWriteDecimal("CFG", "ids_" + LCase(p_ary(i)), All_Power_data.Pins(p_ary(i)).value(Site))
'''            Next i
'''        Next Site
'''    End If
'''    '===========================================================================
    '-------------------------------------------AP & RF
    
    
    If DisableClock Then
        If DisableFRCPinName <> "" And FRC_RelayPin <> "" Then
            Call Enable_FRC_Pins(DisableFRCPinName, FRC_RelayPin, DisableClock)
        Else
            Call Enable_FRC(DisableClockPortName, DisableClock)
        End If
        'If DebugFlag = True Then TheExec.Datalog.WriteComment "print: nWire connect, pin " & DisableClockPortName
    Else
    End If
    
    '-------------------------------------------
    thehdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
    thehdw.Digital.Patgen.HaltWait
    
     'Combine DCVI & DCVS pins for IDS Mapping function 20240624
    If Fuse_Enable Then
        All_Power_Pin = CombineStringList(All_Power_Pin, DCVS_Power_Pin)
        All_Power_Pin = CombineStringList(All_Power_Pin, DCVI_Power_Pin)
        All_Power_Pin = CombineStringList(All_Power_Pin, DCVS_OtherPower_Pin)
    End If
    
        ''20210930 move the IDS calcuation location for validation report
    If DictName <> "" Then
        Call AddStoredMeasurement(DictName, All_Power_data)
    Else
    End If
    
    If Fuse_Enable Then
        Call IDS_Store2Dic_Mapping(All_Power_Pin, All_Power_data, patt)   'If Fuse_StoreName <> "" Then Call IDS_Store2Dic(Fuse_StoreName, DCVS_Power_Pin, All_Power_data, patt)
    Else
    End If
    
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(sTempCalc_Eqn)
    Else
    End If
    
    If FlagWait = True Then
        Call HardIP_WriteFuncResult
    Else
    End If

    'TheHdw.Digital.Patgen.HaltWait
    If EnableDisconnectPins Then
        s_DebugInfo(0) = "***** List IDS Disconnect Pin Start ******"
        s_DebugInfo(1) = "  Pattern Name = " & patt
        s_DebugInfo(2) = "  NC pins = " & sOutput_NCPin
        s_DebugInfo(3) = "  Disconnect pins = " & sOutPut_DisconnectPin
        s_DebugInfo(4) = "  Exclude pins = " & ExcludeDisconnectPins
        s_DebugInfo(5) = "  IllegalExclude pins = " & sOutput_IgnorePin
        s_DebugInfo(6) = "***** List IDS Disconnect Pin End ******"
        s_DebugInfoString = Join(s_DebugInfo, ";")
        DebugPrintFunc patt.value, s_OtherPrintInfo:=s_DebugInfoString
    Else
        DebugPrintFunc patt.value
    End If

    '-------------------------------------------restore pre_measure
    If Interpose_Meas_after <> "" Then
        Call SetForceCondition(Interpose_Meas_after)
    ElseIf Interpose_Meas_before <> "" Then
        Call SetForceCondition("RESTOREPREMEAS")
    Else
    
    End If
    '-------------------------------------------
    'TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
    
    '==== Init pin (restore IFold limit) ====
    thehdw.Digital.ApplyLevelsTiming False, True, False, tlPowered    'SEC DRAM
    '==== Init pin (restore IFold limit) ====

    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, VBT_LIB_DC_IDS, sFuncName)
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function IDS_Delta_calc(Current1_dic As String, Current2_dic As String, power_rail As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim Current_1 As New PinListData
    Dim Current_2 As New PinListData
    Dim Current_delta As New PinListData
    
    Current_1 = GetStoredMeasurement(Current1_dic)
    Current_2 = GetStoredMeasurement(Current2_dic)
    
    Dim p As Variant
    For Each p In Split(power_rail, ",")
        Current_delta.AddPin (p)
        Current_delta.Pins(p).value = Current_1.Pins(p).Subtract(Current_2.Pins(p))
    Next p
    
    Dim i As Long
    Dim site As Variant
    Dim Tname As String
    For i = 0 To Current_delta.Pins.Count - 1
        If InStr(power_rail, Current_delta.Pins.item(i)) <> 0 Then
            Tname = Report_TName_From_Instance("CalcI", Current_delta.Pins.item(i), Current1_dic & "-" & Current2_dic, , , , , , tlForceNone)
            TheExec.Flow.TestLimit Current_delta.Pins(i), Tname:=Tname, PinName:=Current_delta.Pins.item(i), ForceResults:=tlForceFlow, lowVal:=0
        End If
    Next i
        
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_IDS", "IDS_Delta_calc") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function DCVS_IDS_main_current_ratio(Delta_Pin As String, FuseType As String, IdsPinName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim i As Long, j As Long

    Dim IDS_from_Efuse As New SiteDouble
    Dim IDS_from_DCVS As New SiteDouble
    Dim IDS_ratio As New SiteDouble
    Dim IDS_Delta As New SiteDouble
    Dim IDS_PwrName As String
    Dim PwrName As String
    Dim HiLimit_IDS_Delta As Double
    Dim LoLimit_IDS_Delta As Double
    Dim DomainPowerMeasurement_real As New PinListData
    
    ''Seperate the Delta_Pin list into Array, Carter - 20190115, Start
    
    Dim Pins() As String, Pin_Cnt As Long
    Dim Instance_name_str As String: Instance_name_str = UCase(TheExec.DataManager.InstanceName)
    Dim Dic_IDS_fuse_name As New Scripting.Dictionary
    Dim site As Variant 'Carter, 20240304
    TheExec.DataManager.DecomposePinList Delta_Pin, Pins, Pin_Cnt
    
    Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_25C

    'IDS_from_DCVS = 0 'initial
    If FuseType = "CFG" Then
        
        ''====20201230 add for efuse new code====
        Dim opbank As New eFuseBdfBank
        Dim field As New eFuseBdfField
        Dim fieldStr As Variant
        Set opbank = GetBdfBank(FuseType)
        
        For j = 0 To UBound(Pins)
            For Each fieldStr In opbank.Fields
                Set field = opbank.Fields(fieldStr)
    
                If field.Algorithm = alg_ids Then
    
                    If LCase(field.name) = Dic_IDS_fuse_name(Pins(j)) Then
                        IDS_PwrName = Pins(j)
    
                        For Each site In TheExec.sites
                            IDS_from_DCVS = All_Power_data_IDS_GB.Pins(IDS_PwrName).value
                        Next site
                            
                            IDS_from_Efuse = field.DsscDecValue.Multiply(field.Resolution * 0.001)
                            IDS_ratio = IDS_from_DCVS.divide(IDS_from_Efuse)
    
                            TheExec.Flow.TestLimit resultVal:=IDS_from_Efuse, Tname:=IDS_PwrName & "_CP1", PinName:=IDS_PwrName & "_CP1", ForceResults:=tlForceFlow
                            TheExec.Flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_CP2", PinName:=IDS_PwrName & "_CP2", ForceResults:=tlForceFlow
                            TheExec.Flow.TestLimit IDS_ratio, Tname:=IDS_PwrName & "_Ratio", PinName:=IDS_PwrName & "_Ratio", ForceResults:=tlForceFlow
    
                    End If
    
                End If
            Next
        Next j
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_IDS", "DCVS_IDS_main_current_ratio") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''20220124
' [20230908][T-BraC][CC] Use GetStoredData to get SiteDouble
Public Function IDS_main_current_MathFunc(Calc_Pin As PinList, FuseType As String, _
                                            Optional MathFunction As String = "+", _
                                            Optional MathInverse As Boolean = False, _
                                            Optional FusedStage As String, _
                                            Optional Calc_Eqn As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'long
    Dim j As Long
    Dim Pin_Cnt As Long
    'double
    Dim IDS_from_Efuse As New SiteDouble
    Dim IDS_from_DCVS As New SiteDouble
    Dim IDS_Math As New SiteDouble
    'string
    Dim IDS_PwrName As String
    Dim tmp() As String
    Dim tmp2() As String
    Dim SumPin As String
    Dim SplitBySemi() As String
    Dim Pins() As String
    'boolean
    Dim FlagChk As New SiteBoolean
    Dim FlagBinOut As New SiteBoolean: FlagBinOut = True
    'other
    Dim opbank As New eFuseBdfBank
    Dim field As New eFuseBdfField
    Dim CurrentJob_IDSInfo As IDS_Mapping_Info
    Dim SumField As New Dictionary
    Dim site As Variant 'Carter, 20240304
    ''Seperate the Delta_Pin list into Array, Carter - 20190115, Start
    TheExec.DataManager.DecomposePinList Calc_Pin, Pins(), Pin_Cnt

    If FuseType = "CFG" Then
        ''-----------------------------------------------------------------------------------------------------------------------------
        ''equation format
        ''*****************************************************************************************************************************
        ''Alg::Calc_IDSorEFUSE_Sum(Type@PinNameA#FieldName1+FieldName2+...);Calc_IDSorEFUSE_Sum(Type@PinNameB#FieldName5+FieldName6+...)
        ''*****************************************************************************************************************************
        ''Type@
        ''      EFUSE   - sum fused ids value
        ''      IDS     - sum measured ids value
        ''EFUSE type
        ''  PinName#        - the pin is measured ids of pin
        ''  Sum FieldName+  - sum fused ids value
        ''IDS type
        ''  Sum PinName+    - sum measured ids value
        ''
        ''ex - Alg::Calc_Sum(EFUSE@VDD_SOC#IDS_VDD_SOC+IDS_VDD_AVE);Alg::Calc_Sum(IDS@VDD_SOC+VDD_AVE)
        ''-----------------------------------------------------------------------------------------------------------------------------
        If Calc_Eqn <> "" Then
            SumField.RemoveAll
            Call ProcessCalcEquation(Calc_Eqn)
            SplitBySemi = Split(Calc_Eqn, ";")
            For j = 0 To UBound(SplitBySemi)
                tmp = Split(SplitBySemi(j), "@")
                tmp2 = Split(tmp(1), "#")
                SumPin = tmp2(0)
                SumField.Add SumPin, j
            Next
        End If
    
        ''----------------------------------
        ''Check job and get IDS_Mapping_Info
        ''----------------------------------
        For j = 0 To UBound(IDS_MAPPING)
            If UCase(currentJobName) = UCase(IDS_MAPPING(j).Stage) Then
                CurrentJob_IDSInfo = IDS_MAPPING(j)
                Exit For
            End If
        Next j
        
        Set opbank = GetBdfBank(FuseType)
        For j = 0 To UBound(Pins)
            ''link pin name with field name.
            IDS_PwrName = CurrentJob_IDSInfo.MappingDict.item(Pins(j))
            Set field = opbank.Fields(IDS_PwrName)
            
            ''get measured value and fused value
            IDS_from_DCVS = All_Power_data_IDS_GB.Pins(Pins(j))
            If SumField.Exists(Pins(j)) Then
                'get sum fused ids value
                IDS_from_Efuse = GetStoredData(Pins(j) + "_EFUSE")
            Else
                If TheExec.TesterMode = testModeOffline Then
                    IDS_from_Efuse = IDS_Math.Add(0.01 + Rnd() * 0.0001)
                Else
                    IDS_from_Efuse = field.DsscDecValue.Multiply(field.Resolution * 0.001)  '20210406 Modify for new Efuse
                End If
            End If
            
            

            TheExec.Flow.TestLimit resultVal:=IDS_from_Efuse, lowVal:=0.0001, Tname:=IDS_PwrName & "_" & FusedStage, PinName:=Pins(j) & "_" & FusedStage, ForceResults:=tlForceNone
            TheExec.Flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_" & currentJobName, PinName:=Pins(j) & "_" & currentJobName, ForceResults:=tlForceNone
            
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment _
                "site(" & site & ") FusedStage = " & FusedStage & ", eFuseFieldName = " & IDS_PwrName & ", FusedValue = " & Format(IDS_from_Efuse * 1000, "0.000000") & "mA" & _
                ", MeasuredStage = " & currentJobName & ", PinName = " & Pins(j) & ", MeasuredValue = " & Format(IDS_from_DCVS * 1000, "0.000000") & "mA"
            Next

            Select Case (MathFunction)
                Case "+":
                    IDS_Math = IDS_from_DCVS.Add(IDS_from_Efuse)
                    TheExec.Flow.TestLimit IDS_Math, Tname:=IDS_PwrName & "_Add", PinName:=Pins(j) & "_Add", ForceResults:=tlForceFlow
                Case "-":
                    IDS_Math = IIf(Not MathInverse, IDS_from_DCVS.Subtract(IDS_from_Efuse), IDS_from_Efuse.Subtract(IDS_from_DCVS))
                    TheExec.Flow.TestLimit IDS_Math, Tname:=IDS_PwrName & "_Delta", PinName:=Pins(j) & "_Delta", ForceResults:=tlForceFlow
                Case "*":
                    IDS_Math = IIf(Not MathInverse, IDS_from_DCVS.Multiply(IDS_from_Efuse), IDS_from_Efuse.Multiply(IDS_from_DCVS))
                    TheExec.Flow.TestLimit IDS_Math, Tname:=IDS_PwrName & "_Mul", PinName:=Pins(j) & "_Mul", ForceResults:=tlForceFlow, unit:=unitNone
                Case "/":
                    IDS_Math = 9999
                    ''avoid to divide "0".
                    If Not MathInverse Then
                        FlagChk = IDS_from_Efuse.compare(NotEqualTo, 0)
                    Else
                        FlagChk = IDS_from_DCVS.compare(NotEqualTo, 0)
                    End If
                    TheExec.sites.Selected = FlagChk
                    IDS_Math = IIf(Not MathInverse, IDS_from_DCVS.divide(IDS_from_Efuse), IDS_from_Efuse.divide(IDS_from_DCVS))
                    TheExec.sites.Selected = True
                    TheExec.Flow.TestLimit IDS_Math, Tname:=IDS_PwrName & "_Ratio", PinName:=Pins(j) & "_Ratio", ForceResults:=tlForceFlow, unit:=unitNone
                    For Each site In TheExec.sites
                        If FlagChk = False Then
                            TheExec.Datalog.WriteComment "<ERROR> Site(" & site & ") the denominator of the ratio is 0."
                            FlagBinOut = False
                        End If
                    Next
                Case Else:
                    Exit For
            End Select
        Next j
    End If
    
    'bin out the site is denominator = 0
    For Each site In TheExec.sites
        If FlagBinOut = False Then
            TheExec.sites.item(site).result = tlResultFail
        End If
    Next
'============================================================================================
'=  Record Delta IDS to HardKeyReg (added on 2017/7/10)                                     =
'============================================================================================
'    VBT_IEDA_Registry "WLFT_Delta_IDS_PCPU", True
'    VBT_IEDA_Registry "WLFT_Delta_IDS_ECPU", True
'    VBT_IEDA_Registry "WLFT_Delta_IDS_GPU", True
'    VBT_IEDA_Registry "WLFT_Delta_IDS_DCS_DDR", True
'    VBT_IEDA_Registry "WLFT_Delta_IDS_CPU_SRAM", True
'    VBT_IEDA_Registry "WLFT_Delta_IDS_AVE", True
   
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_IDS", "IDS_main_current_MathFunc") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231108][All][Tank] Add Check efuse sheet boolean
Public Function Parse_IDS_Mapping_Table() As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Call ParseIDSMappingTable(True)
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_IDS", "Parse_IDS_Mapping_Table") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
