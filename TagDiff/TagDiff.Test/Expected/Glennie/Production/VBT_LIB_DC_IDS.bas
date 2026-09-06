Attribute VB_Name = "VBT_LIB_DC_IDS"
Option Explicit
'Revision History:
'V0.0 initial bring up

Public gl_All_Power_data_IDS_CP1 As New PinListData
Private Const VBT_LIB_DC_IDS = "VBT_LIB_DC_IDS"

'IDS meas pinlistdata
Public IDS_meas As New PinListData

'=======================20220905_IDS VDD_PCPU Judge > 1.0X=======================
Public All_Power_data_IDS_GB_CORR As New PinListData
Public ids_range_ary_GB() As AutoRange_Info

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
Public gl_IDS_INFO_AllCore_Dic As New Scripting.Dictionary
Public gl_IDS_INFO_4Core_Dic As New Scripting.Dictionary
Public gl_Dic_IDS_fuse_name_AllCore_25C As New Scripting.Dictionary
Public gl_Dic_IDS_fuse_name_9Core_25C As New Scripting.Dictionary
Public gl_Dic_IDS_fuse_name_AllCore_85C As New Scripting.Dictionary
Public gl_Dic_IDS_fuse_name_9Core_105C As New Scripting.Dictionary
Public gl_Dic_IDS_pin_name As New Scripting.Dictionary

Public gb_IDSLimit_Speical_Handle As Boolean
Public gDict_IDSLimit_Special_Handle As New Scripting.Dictionary

Type IDS_INFO
    pin As String
    pat As String
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
                                            Optional nSpecific_Product_Identifier As String = "999", _
                                            Optional Flag As String, Optional dic_loleakpin As Dictionary)    'FixCurrentRange 20220627
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
    Set IDS_VDD_AVE = New SiteVariant   'from Tahiti
    Set IDS_VDD_CPU_SRAM = New SiteVariant
    Set IDS_VDD_DCS_DDR = New SiteVariant
    Set IDS_VDD_DISP = New SiteVariant
    Set IDS_VDD_ECPU = New SiteVariant
    Set IDS_VDD_FIXED = New SiteVariant
    Set IDS_VDD_GPU = New SiteVariant
    Set IDS_VDD_LOW = New SiteVariant
    Set IDS_VDD_PCPU = New SiteVariant
    Set IDS_VDD_SOC = New SiteVariant
    Set IDS_VDD_SRAM_GPU = New SiteVariant
    Set IDS_VDD_SRAM_SOC = New SiteVariant
    Call TheExec.flow.GetTestLimits(FlowLimitsInfo)
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

        ids_range_ary(i).Init_CurrentRange = TheHdw.DCVS.pins(Pin_Ary(i)).CurrentRange.value
        
        ids_range_ary(i).hiLimit = Compare_ForceVal_BV(Pin_Ary(i), Val_Hi(i), True)  'if pins without limit use bincut spec IDSmax as current range
        ''[20240530]remove low limit
        'ids_range_ary(i).loLimit = Compare_ForceVal_BV(Pin_Ary(i), Val_Lo(i), False) 'if pins without limit use bincut spec 0.1IDSmax as current range
        
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
            val = TheHdw.DCVS.pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.level.value    'InitIRange_Setup ids_range_ary(i), WaitTime, SattleTime
        Else
            If ids_range_ary(i).hiLimit = 0 Then
                val = ids_range_ary(i).Init_CurrentRange
            Else
                val = Abs(ids_range_ary(i).hiLimit)  'if pins without bincut spec IDSmax, use current range
            End If
            
            If Pin_Ary(i) Like "VDD_SOC" Then
            
            '    val = thehdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.level.value
                val = 1
            End If
            
            
        End If
        
        ''[240523 min current range to 200uA]
        Call SetCurrentRange(LCase(Pin_Ary(i)), val, WaitTime, ids_range_ary(i).Init_step)
            
    Next i
''**** End - Set init IRange ****
    
    A_HexVS = Split(p_hexvs, ",")
    A_UVS = Split(p_uvs, ",")
    A_VSM = Split(p_vsm, ",")
    
    '''-----------------UFP-----------------
    A_UFP = Split(p_ufp, ",")
    '''-----------------UFP-----------------
    
    TheHdw.Wait 0.01 'add 10ms.    '20230704 code review no need
    TheHdw.Wait WaitTime
    
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        
    Else
        If p_hexvs <> "" Then
            If TheHdw.Alarms.Check Then
                HexVS_Power_data = TheHdw.DCVS.pins(p_hexvs).Meter.Read(tlStrobe, 2000, , tlDCVSMeterReadingFormatAverage)
            Else
                HexVS_Power_data = TheHdw.DCVS.pins(p_hexvs).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
            End If
        End If
        If p_uvs <> "" Then UVS_Power_data = TheHdw.DCVS.pins(p_uvs).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
        If p_vsm <> "" Then VSM_Power_data = TheHdw.DCVS.pins(p_vsm).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        If p_ufp <> "" Then UFP_Power_data = TheHdw.DCVS.pins(p_ufp).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
    
    End If
        
    '-------------------------------------------debug print
    ''                debug_print_pins = AutoRange_Pin
    If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False And gl_Disable_IDS_AutoRange_log = False Then
    If debug_print_pins <> "" Then
        For i = 0 To UBound(Pin_Ary)
            If InStr(LCase(debug_print_pins), LCase(ids_range_ary(i).PinName)) > 0 Then
                s_InstrumentName = UCase(ids_range_ary(i).ChanMapType)
                For Each site In TheExec.sites
                    Select Case s_InstrumentName
                        Case glbConstIns_HEXVS: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0" & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & HexVS_Power_data.pins(ids_range_ary(i).PinName).value(site)
                        Case glbConstIns_VHDVS: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0" & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UVS_Power_data.pins(ids_range_ary(i).PinName).value(site)
                        Case glbConstIns_VSM: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0" & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & VSM_Power_data.pins(ids_range_ary(i).PinName).value(site)
                        Case glbConstIns_VS5A: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0 " & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UFP_Power_data.pins(ids_range_ary(i).PinName).value(site) & " Wait time " & WaitTime * 1000
                        Case glbConstIns_VS800MA: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step 0 " & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UFP_Power_data.pins(ids_range_ary(i).PinName).value(site) & " Wait time " & WaitTime * 1000
                    End Select
                Next site
            End If
        Next i
    End If
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
                            Case glbConstIns_HEXVS: PinVal = HexVS_Power_data_AutoRange.pins(AutoRangePin_Ary(i)).Abs
                            Case glbConstIns_VHDVS: PinVal = UVS256_Power_data_AutoRange.pins(AutoRangePin_Ary(i)).Abs
                            Case glbConstIns_VSM: PinVal = VSM_Power_data_AutoRange.pins(AutoRangePin_Ary(i)).Abs
                            Case glbConstIns_VS5A: PinVal = UFP_Power_data_AutorRange.pins(AutoRangePin_Ary(i)).Abs
                            Case glbConstIns_VS800MA: PinVal = UFP_Power_data_AutorRange.pins(AutoRangePin_Ary(i)).Abs
                        End Select
                        
                        StepNo = ids_range_ary(n_Index).Init_step - j
                        
''                        If AutoRangePin_Ary(i) = "VDD_GPU" Then
''                            If StepNo < 2 Then
''                                StepNo = 2
''                            End If
''                        End If
                        
                        
                        If StepNo >= 0 Then
                            ''[20240529] DCVS AutoRange minimum range set 0.0002
                            If ids_range_ary(n_Index).Range_List(StepNo) > 0.00002 Then
                                ''[20240608] HEXVS AutoRange minimum range set 0.01
                                If s_InstrumentName = glbConstIns_HEXVS And ids_range_ary(n_Index).Range_List(StepNo) < 0.1 Then
                                ElseIf AutoRangePin_Ary(i) = "VDD_GPU" And StepNo < 2 Then
                                    StepNo = 2
                                    
                    
                                ''do nothing
                                Else
                                    DropRngSite = PinVal.compare(LessThan, ids_range_ary(n_Index).Range_List(StepNo) - ids_range_ary(n_Index).Accuracy_List(StepNo + 1))
                                    If DropRngSite.Any(True) Then
                                        Flag_PerPin_autorange = True
                                        TheExec.sites.Selected = DropRngSite
                                        TheHdw.DCVS.pins(AutoRangePin_Ary(i)).SetCurrentRanges ids_range_ary(n_Index).Range_List(StepNo), ids_range_ary(n_Index).Range_List(StepNo)
                                        TheExec.sites.Selected = True
                                        SattleTime = ids_range_ary(n_Index).WaitTime_List(StepNo)
                                        If SattleTime > WaitTime Then WaitTime = SattleTime
''                                        If (LCase(AutoRangePin_Ary(i)) = "vdd_amph_ddr") Or (LCase(AutoRangePin_Ary(i)) = "vddio12_grp") Or (UCase(AutoRangePin_Ary(i)) = "VDDIO12_AOP_2") Then
''
''                                            WaitTime = WaitTime + 0.01
''
''                                        End If
''
''
                                    End If
                                End If
                            ElseIf (LCase(AutoRangePin_Ary(i)) = "vdd_amph_ddr") Or (LCase(AutoRangePin_Ary(i)) = "vddio12_grp") Or (UCase(AutoRangePin_Ary(i)) = "VDDIO12_AOP_2") Then
                                If ids_range_ary(k).Range_List(StepNo + 1) = 0.0002 Then WaitTime = 0.01
                                   If SattleTime > WaitTime Then WaitTime = SattleTime
                                
                             
                            End If
                        End If
                    End If
                    Set PinVal = Nothing
                    Flag_AllPin_autorange = Flag_AllPin_autorange Or Flag_PerPin_autorange
                    Next i
                    If Flag_AllPin_autorange Then
                        
                        TheHdw.Wait WaitTime
                        ''Add measurement points to prevent error 20171011 (M9)
                        If p_hexvs_autorRange <> "" Then
                            If TheHdw.Alarms.Check Then
                                HexVS_Power_data_AutoRange = TheHdw.DCVS.pins(p_hexvs_autorRange).Meter.Read(tlStrobe, 2000, , tlDCVSMeterReadingFormatAverage)
                            Else
                                HexVS_Power_data_AutoRange = TheHdw.DCVS.pins(p_hexvs_autorRange).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                            End If
                        End If
                        If p_uvs256_autorRange <> "" Then UVS256_Power_data_AutoRange = TheHdw.DCVS.pins(p_uvs256_autorRange).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
                        If p_vsm_autorRange <> "" Then VSM_Power_data_AutoRange = TheHdw.DCVS.pins(p_vsm_autorRange).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                        
                        '''-----------------UFP-----------------
                        If p_ufp_autorRange <> "" Then UFP_Power_data_AutorRange = TheHdw.DCVS.pins(p_ufp_autorRange).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                        '''-----------------UFP-----------------
                        
                        '-------------------------------------------debug print
        ''                debug_print_pins = AutoRange_Pin
                        
                        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False And gl_Disable_IDS_AutoRange_log = False Then
                        If debug_print_pins <> "" Then
                            For i = 0 To UBound(AutoRangePin_Ary)
                                If dic_MeasurePowerPinIndex.Exists(LCase(AutoRangePin_Ary(i))) Then     'If InStr(LCase(debug_print_pins), LCase(ids_range_ary(i).PinName)) > 0 Then
                                    s_InstrumentName = UCase(ids_range_ary(i).ChanMapType)
                                    For Each site In TheExec.sites
                                        Select Case s_InstrumentName
                                            Case glbConstIns_HEXVS: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).Meter.CurrentRange.value & ", Current: " & HexVS_Power_data_AutoRange.pins(ids_range_ary(i).PinName).value(site)
                                            Case glbConstIns_VHDVS: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UVS256_Power_data_AutoRange.pins(ids_range_ary(i).PinName).value(site)
                                            Case glbConstIns_VSM: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & VSM_Power_data_AutoRange.pins(ids_range_ary(i).PinName).value(site)
                                            Case glbConstIns_VS5A: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UFP_Power_data_AutorRange.pins(ids_range_ary(i).PinName).value(site)
                                            Case glbConstIns_VS800MA: TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(i).PinName & ", Step " & j & ", Irange: " & TheHdw.DCVS.pins(ids_range_ary(i).PinName).CurrentRange.value & ", Current: " & UFP_Power_data_AutorRange.pins(ids_range_ary(i).PinName).value(site)
                                        End Select
                                    Next site
                                End If
                            Next i
                        End If
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

    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        For i = 0 To UBound(Pin_Ary)
            Power_data.AddPin Pin_Ary(i)
        Next i
        
    Else
        For i = 0 To UBound(A_HexVS)
            Power_data.AddPin (A_HexVS(i))
            If LCase("*," & p_hexvs_autorRange & ",*") Like LCase("*," & A_HexVS(i) & ",*") Then
                Power_data.pins(A_HexVS(i)) = HexVS_Power_data_AutoRange.pins(A_HexVS(i))
            Else
                Power_data.pins(A_HexVS(i)) = HexVS_Power_data.pins(A_HexVS(i))
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
                Power_data.pins(A_UVS(i)) = UVS256_Power_data_AutoRange.pins(A_UVS(i))
            Else
                Power_data.pins(A_UVS(i)) = UVS_Power_data.pins(A_UVS(i))
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
                Power_data.pins(A_VSM(i)) = VSM_Power_data_AutoRange.pins(A_VSM(i))
            Else
                Power_data.pins(A_VSM(i)) = VSM_Power_data.pins(A_VSM(i))
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
                Power_data.pins(A_UFP(i)) = UFP_Power_data_AutorRange.pins(A_UFP(i))
            Else
                Power_data.pins(A_UFP(i)) = UFP_Power_data.pins(A_UFP(i))
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
                        If Flag_BinX_Info_Parsed = True And CurrentPassBinCutNum_normal(site) = 2 Then
                        Else
                            GlbUtility.WriteDlg "site:" & site & ", product_identifier " & CurrentPassBinCutNum_normal(site) & " > Total_Bincut_Num " & Total_Bincut_Num & " , Error!!!"
                            TheExec.flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="product_identifier Error"
                        End If
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
            VMain = Format(TheHdw.DCVS.pins(Power_data.pins(CorePowerPin_Ary(i))).Voltage.Main.value, "0.00")
            Hilimit_IDS = 0
            Lolimit_IDS = 0
            If dic_MeasurePowerPinIndex.Exists(LCase(CorePowerPin_Ary(i))) Then
                n_Index = dic_MeasurePowerPinIndex(LCase(CorePowerPin_Ary(i)))

                For Each site In TheExec.sites
                    ''[240530]remove low limit
                    Hilimit_IDS = Compare_ForceVal_BV(Pin_Ary(n_Index), Val_Hi(n_Index), True, isUse_Product_Identifier, nTempBinCutNum(site))
                    'Lolimit_IDS = Compare_ForceVal_BV(Pin_Ary(n_Index), Val_Lo(n_Index), True, isUse_Product_Identifier, nTempBinCutNum(site))
                    If TheExec.TesterMode = testModeOffline Or gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
                        d_SimulationValue = (Hilimit_IDS + Lolimit_IDS) / 2
                        Power_data.pins(Pin_Ary(n_Index)).value(site) = d_SimulationValue
                    End If
                    
                    If ENG_Limit = False Then
                        TheExec.Datalog.WriteComment "Site(" & site & ") power pin : " & CorePowerPin_Ary(i) & " - value : " & Format((Power_data.pins(CorePowerPin_Ary(i)).value(site) * 1000), "0.000000") & " mA"
                    ElseIf (TheExec.sites.item(site).FlagState(Flag) = logicTrue And Not (TheExec.sites.item(site).FlagState("F_IDS_CP1_HIGH_LEAKAGE") = logicTrue)) And (dic_loleakpin.Exists(UCase(CorePowerPin_Ary(i)))) Then
                        TheExec.flow.TestLimit resultVal:=Power_data.pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, hiVal:=dic_loleakpin(UCase(CorePowerPin_Ary(i)))
                    Else
                    
                        If Hilimit_IDS = 0 Then     'If Hilimit_IDS = 0 And Lolimit_IDS = 0 Then
                            TheExec.flow.TestLimit resultVal:=Power_data.pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, highCompareSign:=tlSignNone       ', highCompareSign:=tlSignNone, lowCompareSign:=tlSignNone
                        Else
                            TheExec.flow.TestLimit resultVal:=Power_data.pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, hiVal:=Hilimit_IDS   'theexec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, hiVal:=Hilimit_IDS, lowVal:=Lolimit_IDS
                        End If
                        
                    End If
                    TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
                Next site

            End If
            
            TheExec.Datalog.WriteComment "Current I Range: " & CorePowerPin_Ary(i) & "--->" & TheHdw.DCVS.pins(CorePowerPin_Ary(i)).Meter.CurrentRange.value
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1

        End If
        If UCase(TheExec.DataManager.instancename) Like "IDS_IDS_PP*" Then
           
        
            Dim m_valStr As String
            Dim m_site As Variant
     
            For Each m_site In TheExec.sites
              
                m_valStr = Format(Power_data.pins(CorePowerPin_Ary(i)).value(m_site) * 1000, "0.000000")
    
                If UCase(CorePowerPin_Ary(i)) = "VDD_AVE" Then
                    IDS_VDD_AVE = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_CPU_SRAM" Then
                    IDS_VDD_CPU_SRAM = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_DCS_DDR" Then
                    IDS_VDD_DCS_DDR = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_DISP" Then
                    IDS_VDD_DISP = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_ECPU" Then
                    IDS_VDD_ECPU = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) Like "VDD_FIXED*" Then
                    IDS_VDD_FIXED = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_GPU" Then
                    IDS_VDD_GPU = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) Like "VDD_LOW*" Then
                    IDS_VDD_LOW = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_PCPU" Then
                    IDS_VDD_PCPU = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_SOC" Then
                    IDS_VDD_SOC = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_SRAM_GPU" Then
                    IDS_VDD_SRAM_GPU = m_valStr
                ElseIf UCase(CorePowerPin_Ary(i)) = "VDD_SRAM_SOC" Then
                    IDS_VDD_SRAM_SOC = m_valStr
                End If
    
            Next m_site
        End If
    Next j: Next i

    For k = 0 To OtherPowerPin_Cnt - 1
        If gl_GetInstrument_Dic.Exists(LCase(OtherPowerPin_Ary(k))) Then

            'Tname = theexec.DataManager.instanceName & "_" & OtherPowerPin_Ary(k)        'add pin name Aruba 2017/12/28
            Tname = Report_TName_From_Instance("I", OtherPowerPin_Ary(k), , , , , , , tlForceNone)
            VMain = Format(TheHdw.DCVS.pins(Power_data.pins(OtherPowerPin_Ary(k))).Voltage.Main.value, "0.00")

            Hilimit_IDS = 0
            Lolimit_IDS = 0
            
            If dic_MeasurePowerPinIndex.Exists(LCase(OtherPowerPin_Ary(k))) Then
                n_Index = dic_MeasurePowerPinIndex(LCase(OtherPowerPin_Ary(k)))

                For Each site In TheExec.sites
                    Hilimit_IDS = Compare_ForceVal_BV(Pin_Ary(n_Index), Val_Hi(n_Index), True, isUse_Product_Identifier, nTempBinCutNum(site))
                    ''[240530]remove low limit
                    'Lolimit_IDS = Compare_ForceVal_BV(Pin_Ary(n_Index), Val_Lo(n_Index), True, isUse_Product_Identifier, nTempBinCutNum(site))
                    If TheExec.TesterMode = testModeOffline Or gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
                        d_SimulationValue = (Hilimit_IDS + Lolimit_IDS) / 2
                        Power_data.pins(Pin_Ary(n_Index)).value(site) = d_SimulationValue
                    End If
                    
                    If ENG_Limit = False Then
                        TheExec.Datalog.WriteComment "Site(" & site & ") power pin : " & OtherPowerPin_Ary(k) & " - value : " & Format((Power_data.pins(OtherPowerPin_Ary(k)).value(site) * 1000), "0.000000") & " mA"
                    Else
                    
                        If Hilimit_IDS = 0 Then     'If Hilimit_IDS = 0 And Lolimit_IDS = 0 Then
                            TheExec.flow.TestLimit resultVal:=Power_data.pins(OtherPowerPin_Ary(k)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, highCompareSign:=tlSignNone
                        Else
                            TheExec.flow.TestLimit resultVal:=Power_data.pins(OtherPowerPin_Ary(k)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, hiVal:=Hilimit_IDS
                        End If
                        
                    End If
                    TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
                Next site

            End If
            
            TheExec.Datalog.WriteComment "Current I Range: " & OtherPowerPin_Ary(k) & "--->" & TheHdw.DCVS.pins(OtherPowerPin_Ary(k)).Meter.CurrentRange.value
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1

        End If
    Next k

    'recover range setup
    
    'Pin_Cnt = CorePowerPin_Cnt + OtherPowerPin_Cnt
    
    For i = 0 To UBound(Pin_Ary)
        TheHdw.DCVS.pins(Pin_Ary(i)).SetCurrentRanges ids_range_ary(i).Init_CurrentRange, ids_range_ary(i).Init_CurrentRange
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
                                                         Optional AutoRange_Pin As String, Optional Flag As String, Optional dic_loleakpin As Dictionary)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
gl_Disable_IDS_AutoRange_log = True
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
    Call TheExec.flow.GetTestLimits(FlowLimitsInfo)
    
    'if no Use-Limits on this test, FlowLimitsInfo is nothing
    If FlowLimitsInfo Is Nothing Then
        If isDebugMode Then TheExec.AddOutput "Could not get the limits info", vbRed, True
        Exit Function
    End If

    Dim Val_Hi() As String
    Dim Val_Lo() As String
    FlowLimitsInfo.GetHighLimits Val_Hi
    FlowLimitsInfo.GetLowLimits Val_Lo
    
    Dim dcvs_pincnt As Long
    dcvs_pincnt = Power_data.pins.Count

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
    
    TheExec.DataManager.DecomposePinList CorePower_Pin, CorePowerPin_Ary, CorePowerPin_Cnt

    For i = 0 To CorePowerPin_Cnt - 1
        If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) <> "N/C" Then All_Power_Pin = All_Power_Pin & "," & CorePowerPin_Ary(i)
    Next i

    If All_Power_Pin <> "" Then All_Power_Pin = right(All_Power_Pin, Len(All_Power_Pin) - 1)

    Pin_Ary = Split(All_Power_Pin, ",")
    ReDim IDS_ini_Current_range(UBound(Pin_Ary)) As Double
    WaitTime = 100 * us
    
    ''**** Start - Get PowerPin Info ****
    Dim ids_range_ary() As AutoRange_Info
    ReDim ids_range_ary(UBound(Pin_Ary))
    For i = 0 To UBound(Pin_Ary)
        For j = 0 To UBound(PowerPin_range_ary)
            If Pin_Ary(i) = PowerPin_range_ary(j).PinName Then
                ids_range_ary(i) = PowerPin_range_ary(j)
            
            End If
        Next j
    Next i
    ''**** End - Get PowerPin Info ****

    ' Set init IRange
    For i = 0 To UBound(Pin_Ary)
    
        ids_range_ary(i).Init_CurrentRange = TheHdw.DCVI.pins(Pin_Ary(i)).CurrentRange.value
        ids_range_ary(i).Init_Source_FoldLimit = TheHdw.DCVI.pins(Pin_Ary(i)).Current
        ids_range_ary(i).hiLimit = Compare_ForceVal_BV(Pin_Ary(i), Val_Hi(i + dcvs_pincnt)) 'if pins without limit use bincut spec IDSmax as current range
    
        If LCase(ids_range_ary(i).ChanMapType) = "dc-07" Then
            p_uvi80 = p_uvi80 & "," & Pin_Ary(i)
            TheHdw.DCVI.pins(Pin_Ary(i)).Meter.mode = tlDCVIMeterCurrent '20180115 Rick
        End If
        
        If (FlowLimitForInitIRange = False) Then
            InitIRange_Setup ids_range_ary(i), WaitTime, SattleTime
        
        Else
            If ids_range_ary(i).hiLimit = 0 Then val = ids_range_ary(i).Init_CurrentRange Else val = Abs(ids_range_ary(i).hiLimit)  'if pins without bincut spec IDSmax, use current range
'                If Val_Hi(i) = "" Then Val = range_ary(i).Init_CurrentRange Else Val = Abs(Val_Hi(i)) 'if pins without limit use current range
                'If TheExec.Flow.EnableWord("Temp_25C") Or TheExec.Flow.EnableWord("Temp_N25C") Then
                    
                    ''[240523 min current range to 200uA]
                    If val < 0.0002 Then val = 0.0002
                    
                    For j = 0 To UBound(ids_range_ary(i).Range_List)

                        If val <= ids_range_ary(i).Range_List(j) Then
                            ids_range_ary(i).Init_step = j
                            SattleTime = ids_range_ary(i).WaitTime_List(j)
                            'If UCase(Pin_Ary(i)) Like "VDD_SRAM" Or UCase(Pin_Ary(i)) Like "VDD_SOC" Then
                            '    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange 20 * ma, 20 * ma 'ids_range_ary(i).Range_List(j), ids_range_ary(i).Range_List(j)
                            '    TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.value = 20 * ma 'ids_range_ary(i).Range_List(j)
                            'Else
                                TheHdw.DCVI.pins(Pin_Ary(i)).SetCurrentAndRange ids_range_ary(i).Range_List(j), ids_range_ary(i).Range_List(j)
                                TheHdw.DCVI.pins(Pin_Ary(i)).CurrentRange.value = ids_range_ary(i).Range_List(j)
                            'End If
        
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                            Exit For
                        End If
                    Next j
                'ElseIf TheExec.Flow.EnableWord("Temp_85C") Then
    ''                TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange ids_range_ary(i).Range_List(UBound(ids_range_ary(i).Range_List)), ids_range_ary(i).Range_List(UBound(ids_range_ary(i).Range_List))
    ''                TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.Value = ids_range_ary(i).Range_List(UBound(ids_range_ary(i).Range_List))
                '    ids_range_ary(i).Init_step = UBound(ids_range_ary(i).Range_List)
                '    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange 200 * ma, 200 * ma 'ids_range_ary(i).Range_List(j), ids_range_ary(i).Range_List(j)
                '    TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.value = 200 * ma 'ids_range_ary(i).Range_List(j)
                'End If
        End If
'        Debug.Print i & ":" & CorePower_Pin_Ary(i)
    Next i

    If p_uvi80 <> "" Then p_uvi80 = right(p_uvi80, Len(p_uvi80) - 1)
    A_UVI80 = Split(p_uvi80, ",")

    TheHdw.Wait 0.035 'add 10ms.
    TheHdw.Wait WaitTime

    ReDim AutoRangePin(UBound(CorePowerPin_Ary))
    AutoRangePin_Ary = CorePowerPin_Ary

    If p_uvi80 <> "" Then UVI80_Power_data = TheHdw.DCVI.pins(p_uvi80).Meter.Read(tlStrobe, 1, , tlDCVIMeterReadingFormatAverage)
    
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
                                    TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(k).PinName & ", Step 0" & ", Irange: " & TheHdw.DCVI.pins(ids_range_ary(k).PinName).Meter.CurrentRange.value & ", Current: " & UVI80_Power_data.pins(ids_range_ary(k).PinName).value(site)
                                End If
                            Next site
                        End If
                    End If
                Next k
            Next i
        End If
    End If
    '-------------------------------------------
    
    Dim Stop_Step, StepNo As Integer
    If Search_Step = "" Then
        Stop_Step = 7
    Else
        Stop_Step = CLng(Search_Step)
    End If
    '========================================================================================auto range search
    Dim p_uvi80_autorRange As String
    
    Dim A_uvi80_autorRange() As String
    
    Dim UVI80_Power_data_AutoRange As New PinListData
    
    Dim Flag_PerPin_autorange As Boolean
    Dim Flag_AllPin_autorange As Boolean
    
    If TheExec.flow.enableWord("CurrentProfile") Or TheExec.flow.enableWord("VoltageProfile") Then
    
    Else
        ''****Start -  Set Auto IRange ****
        If AutoRange_Pin <> "" Then
        
            TheExec.DataManager.DecomposePinList AutoRange_Pin, AutoRangePin_Ary, AutoRangePin_Cnt
    
            For i = 0 To UBound(AutoRangePin_Ary)
                For k = 0 To UBound(ids_range_ary)
                    If LCase(AutoRangePin_Ary(i)) = LCase(ids_range_ary(k).PinName) Then
                        If LCase(ids_range_ary(k).ChanMapType) = "dc-07" Then
                            p_uvi80_autorRange = p_uvi80_autorRange & "," & AutoRangePin_Ary(i)
                        End If
                        Exit For
                    End If
                Next k
            Next i
    
            If p_uvi80_autorRange <> "" Then
                p_uvi80_autorRange = right(p_uvi80_autorRange, Len(p_uvi80_autorRange) - 1)
                A_uvi80_autorRange = Split(p_uvi80_autorRange, ",")
                UVI80_Power_data_AutoRange = UVI80_Power_data
            End If

            For j = 1 To Stop_Step
    
                WaitTime = 260 * us
    
                For i = 0 To UBound(AutoRangePin_Ary)
    
                    Flag_PerPin_autorange = False
    
                    For k = 0 To UBound(ids_range_ary)
    
                        If LCase(AutoRangePin_Ary(i)) = LCase(ids_range_ary(k).PinName) Then
    
                            If LCase(ids_range_ary(k).ChanMapType) = "dc-07" Then
                                PinVal = UVI80_Power_data_AutoRange.pins(AutoRangePin_Ary(i)).Abs

                            End If
                            
                            StepNo = ids_range_ary(k).Init_step - j
                            If StepNo >= 0 Then
                                ''[20240529] AutoRange minimum range set 0.0002
                                If ids_range_ary(k).Range_List(StepNo) > 0.00002 Then ''20230710, If current range is larger than 20uA, AutoRange will be enabled.
                                    'DropRngSite = PinVal.compare(LessThan, ids_range_ary(k).Range_List(StepNo) - ids_range_ary(k).Accuracy_List(StepNo))
                                    DropRngSite = PinVal.compare(LessThan, ids_range_ary(k).Range_List(StepNo) * 0.5) 'KC fix alarm issue 211217
                                    If DropRngSite.Any(True) Then
                                        Flag_PerPin_autorange = True
                                        TheExec.sites.Selected = DropRngSite
                                        If Not (UCase(AutoRangePin_Ary(i)) Like "VDD_SRAM" Or UCase(AutoRangePin_Ary(i)) Like "VDD_SOC") Then
                                            'disable mode alarm for current range change on DCVI power pin. 20211224
                                            TheHdw.DCVI.pins(p_uvi80_autorRange).Alarm(tlDCVIAlarmMode) = tlAlarmOff
                                            TheHdw.DCVI.pins(AutoRangePin_Ary(i)).SetCurrentAndRange ids_range_ary(k).Range_List(StepNo), ids_range_ary(k).Range_List(StepNo)
                                        End If
                                        TheExec.sites.Selected = True
                                        SattleTime = ids_range_ary(k).WaitTime_List(StepNo)
                                        If SattleTime > WaitTime Then WaitTime = SattleTime
                                    End If
                                End If
                            End If
                            
                            Exit For
    
                        End If
    
                    Next k
    
                    Set PinVal = Nothing
    
                    Flag_AllPin_autorange = Flag_AllPin_autorange Or Flag_PerPin_autorange
                Next i
    
                If Flag_AllPin_autorange Then
                    'If TheExec.Flow.EnableWord("Temp_85C") Then WaitTime = WaitTime + 0.015
                    TheHdw.Wait WaitTime

                    'enable mode alarm for current range change on DCVI power pin. 20211224
                    TheHdw.DCVI.pins(p_uvi80_autorRange).Alarm(tlDCVIAlarmMode) = tlAlarmForceFail

                    If p_uvi80_autorRange <> "" Then UVI80_Power_data_AutoRange = TheHdw.DCVI.pins(p_uvi80_autorRange).Meter.Read(tlStrobe, 1, , tlDCVIMeterReadingFormatAverage)
                    
                    '-------------------------------------------debug print, disable with HIP TTR enable word
                    If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False And gl_Disable_IDS_AutoRange_log = False Then
    ''                debug_print_pins = AutoRange_Pin
                        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False And gl_Disable_IDS_AutoRange_log = False Then
                        If debug_print_pins <> "" Then
                            For i = 0 To UBound(AutoRangePin_Ary)
                                For k = 0 To UBound(ids_range_ary)
                                    If UCase(AutoRangePin_Ary(i)) = UCase(ids_range_ary(k).PinName) Then
                                        If InStr(LCase(debug_print_pins), LCase(ids_range_ary(k).PinName)) > 0 Then
                                            For Each site In TheExec.sites
                                                If LCase(ids_range_ary(k).ChanMapType) = "dc-07" Then
                                                    TheExec.Datalog.WriteComment "Site(" & site & "), " & ids_range_ary(k).PinName & ", Step " & j & ", Irange: " & TheHdw.DCVI.pins(ids_range_ary(k).PinName).CurrentRange.value & ", Current: " & UVI80_Power_data_AutoRange.pins(ids_range_ary(k).PinName).value(site)
    
                                                End If
                                            Next site
                                        End If
                                    End If
                                Next k
                            Next i
                        End If
                      End If
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
    '========================================================================================auto range search

    For i = 0 To UBound(A_UVI80)
        Power_data.AddPin (A_UVI80(i))
        If LCase("*," & p_uvi80_autorRange & ",*") Like LCase("*," & A_UVI80(i) & ",*") Then
            Power_data.pins(A_UVI80(i)) = UVI80_Power_data_AutoRange.pins(A_UVI80(i))
        Else
            Power_data.pins(A_UVI80(i)) = UVI80_Power_data.pins(A_UVI80(i))
        End If
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                Power_data.pins(A_UVI80(i)).value(site) = 0.01 + Rnd() * 0.0001
            Next site
        End If
    Next i

    For i = 0 To CorePowerPin_Cnt - 1: For j = 0 To repeat_count - 1
        If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) <> "N/C" Then

            ''20180315 Abel change naming'            Tname = TheExec.DataManager.InstanceName & "_" & j
'            Tname = TheExec.DataManager.instanceName 'No need _0
            Tname = Report_TName_From_Instance("I", Power_data.pins(CorePowerPin_Ary(i)), , , , , , , tlForceNone)
            If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) Like "*DCVI*" Then
                VMain = Format(TheHdw.DCVI.pins(Power_data.pins(CorePowerPin_Ary(i))).Voltage, "0.00")
            End If
            TheExec.Datalog.WriteComment (glb_TestInstance & " =====> Curr_meas Meter I range setting, " & CorePowerPin_Ary(i) & " =" & TheHdw.DCVI.pins(CorePowerPin_Ary(i)).Meter.CurrentRange.value)
            If TPModeAsCharz_GLB = True Then 'wc 180319 for charz
                TheExec.flow.TestLimit resultVal:=Power_data.pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, Tname:=Tname, formatStr:="%.3f", ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                TheExec.Datalog.WriteComment "Current I Range: " & CorePowerPin_Ary(i) & "--->" & TheHdw.DCVI.pins(CorePowerPin_Ary(i)).Meter.CurrentRange.value
            ElseIf (TheExec.sites.item(site).FlagState(Flag) = logicTrue And Not (TheExec.sites.item(site).FlagState("F_IDS_CP1_HIGH_LEAKAGE") = logicTrue)) And (dic_loleakpin.Exists(UCase(CorePowerPin_Ary(i)))) Then
                TheExec.flow.TestLimit resultVal:=Power_data.pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, Tname:=Tname, formatStr:="%.3f", ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow, hiVal:=dic_loleakpin(UCase(CorePowerPin_Ary(i)))
            Else
                Powerpin_log = Replace(UCase(CorePowerPin_Ary(i)), "_", "")
                TheExec.flow.TestLimit resultVal:=Power_data.pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, Tname:=Tname, formatStr:="%.3f", ForceVal:=VMain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow   ''20180315 Abel add power pin name
                TheExec.Datalog.WriteComment "Current I Range: " & CorePowerPin_Ary(i) & "--->" & TheHdw.DCVI.pins(CorePowerPin_Ary(i)).Meter.CurrentRange.value
            End If
        End If
    Next j: Next i

    For i = 0 To UBound(Pin_Ary)
        If TheExec.DataManager.ChannelType(Pin_Ary(i)) Like "*DCVI*" Then
            '''''TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange IDS_ini_Current_range(i), IDS_ini_Current_range(i)
            TheHdw.DCVI.pins(Pin_Ary(i)).SetCurrentAndRange ids_range_ary(i).Init_Source_FoldLimit, ids_range_ary(i).Init_CurrentRange
        End If
    Next i

    TheHdw.Wait 0.003

    Exit Function
    
    
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function IDS_main_current(patt As Pattern, _
                                      DCVS_Power_Pin As String, DCVI_Power_Pin As String, _
                                      DCVS_OtherPower_Pin As String, _
                                      repeat_count As Long, _
                                      FlowLimitForInitIRange As Boolean, _
                             Optional Search_Step As String, _
                             Optional DisableClock As Boolean = False, _
                             Optional FlagWait As Boolean = False, Optional FlagPatFinishMeasure As Boolean = False, _
                             Optional FixCurrentRange As String, _
                             Optional CharInputString As String, Optional BV_SpecialHandle As String, Optional DisableClockPortName As String, Optional DisableFRCPinName As String, Optional FRC_RelayPin As String, _
                             Optional RTOS_Setup As Boolean = False, Optional RTOSCmd As String = vbNullString, Optional RTOSTimeOut As Double, Optional DisconnectClock As Boolean = True, Optional debug_print_pins As String, Optional NotPrintOutLimit As Boolean = False, Optional Interpose_Meas_before As String, Optional Interpose_Meas_after As String, _
                             Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, Optional CUS_Str_DigSrcData As String = vbNullString, _
                             Optional DictName As String, Optional Fuse_Enable As Boolean, Optional Calc_Eqn As String, Optional AutoRange_Pin As String, Optional Validating_ As Boolean, _
                             Optional isUse_Product_Identifier As Boolean = False, Optional sSpecific_Product_Identifier As String = "999", Optional EnableDisconnectPins As Boolean = False, Optional ExcludeDisconnectPins As String = vbNullString, Optional Flag As String, Optional leaklimit As String)    'Carter, 20190315     'FixCurrentRange 20220627       '20221014 Tank add bincut sheet reference product identifier

    
    gl_IDS_INFO_Dic.RemoveAll
    Dim p As Variant
    Dim p_ary() As String
    Dim pinCnt As Long
    Dim MeasCurr_HexVS As New PinListData
    Dim MeasCurr As New PinListData
    Dim MeasCurr_copy As New PinListData
    Dim power_pin As String
    Dim TestNum() As Long, Cnt1 As Long
    Dim i As Long, j As Long
    Dim repeat_judge As Long

    Dim All_Power_data As New PinListData
    Dim site As Variant

    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim CLK_Pins As String

    Dim rtnPatNames() As String
    Dim patcnt As Long
    Dim InDSPwave As New DSPWave
    Dim temp_instance_name As String: temp_instance_name = TheExec.DataManager.instancename

    Dim sTempCalc_Eqn As String

    Dim sOutPut_DisconnectPin As String
    Dim sOutput_IgnorePin As String
    Dim sDontCarePin As String
    Dim sOutput_NCPin As String
    Dim isDoDisconnectPinPass As Boolean
    Dim s_DebugInfo(6) As String
    Dim s_DebugInfoString As String
    
    Dim sFuncName As String:: sFuncName = "IDS_main_current"
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)

    Dim dic_loleakpin As New Dictionary
    Dim sa_loleakPin() As String
    Dim sa_loleakPinName() As String
    Dim da_loleakPinData() As Double
    Dim q As Long
    Dim s_loleakPinName As String
    Dim Hilimit_IDS As Double

    On Error GoTo errHandler
    
    If Validating_ Then 'Carter, 20190315
        Call PrLoadPattern(patt.value)
        Exit Function    ' Exit after validation
    End If
    
    ' ===== 20220113 add for IbizaA0 RTOS  ======
    If patt <> "" Then
    Else
        If RTOSCmd <> "" Then
            RTOS_IDS RTOSCmd, RTOSTimeOut
        End If
    End If
    '===============================================
    
    If isUse_Product_Identifier Then    'replace Calc_Eqn string to add production identifier
        sTempCalc_Eqn = Replace(Calc_Eqn, "~", "~d")    'use efuse value in device
    Else
        If IsNumeric(sSpecific_Product_Identifier) = False Then
            sSpecific_Product_Identifier = "999"
        End If
        sTempCalc_Eqn = Replace(Calc_Eqn, "~", "~" & sSpecific_Product_Identifier)      'use specific production identifier
    End If
    
    If ENG_SweepPin Then
        TheExec.Datalog.WriteComment "[" & glb_TestInstance & "  start]"
        TheExec.Datalog.WriteComment "Threshold_CriticalPin =" & Threshold_CriticalPin
        ENG_Limit = False
        Call FindInputPin(patt, FlagWait, DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData)

        Call SweepPinInPattern(patt, FlagWait, DCVS_Power_Pin, DCVS_OtherPower_Pin, repeat_count, FlowLimitForInitIRange, _
            DisableClock, Search_Step, DisableClockPortName, DisconnectClock, debug_print_pins, AutoRange_Pin, _
            DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, _
            Interpose_Meas_before, Interpose_Meas_after, DisableFRCPinName, FRC_RelayPin)

        TheExec.Datalog.WriteComment "[" & glb_TestInstance & "  end]"
    End If
    ENG_Limit = True
    
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
'    If TheExec.DevChar.Setups.IsRunning And Not glb_ShmooEnable Then Exit Function ' Add for IDS shmoo 230130
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
    End If

''**** For Special Low Leakage Bin1 Device 20230706 ****
'leaklimit example = VDD_PCPU:406.4;VDD_ECPU:131;VDD_GPU:173.8;VDD_SOC:97.608;VDD_DCS_DDR:13.706
If leaklimit <> "" Then
        dic_loleakpin.RemoveAll
        sa_loleakPin = Split(leaklimit, ";")
        ReDim sa_loleakPinName(UBound(sa_loleakPin)) As String
        ReDim da_loleakPinData(UBound(sa_loleakPin)) As Double
        If UBound(sa_loleakPin) = 0 Then ReDim Preserve sa_loleakPin(1)
        For q = 0 To UBound(sa_loleakPin)
            Hilimit_IDS = 0
            sa_loleakPinName(q) = Split(sa_loleakPin(q), ":")(0)
            If UBound(Split(sa_loleakPin(q), ":")) = 1 Then
                If IsNumeric((Split(sa_loleakPin(q), ":")(1))) = True Then
                    da_loleakPinData(q) = (Split(sa_loleakPin(q), ":")(1)) / 1000
                    If dic_loleakpin.Exists(sa_loleakPinName(q)) Then
                    Else
                        If gDict_IDSLimit_Special_Handle.Exists(UCase(sa_loleakPinName(q))) Then
                            Hilimit_IDS = da_loleakPinData(q)
                            Limit_Special_Handle sa_loleakPinName(q), Hilimit_IDS
                            dic_loleakpin.Add UCase(sa_loleakPinName(q)), Hilimit_IDS
                        Else
                            dic_loleakpin.Add UCase(sa_loleakPinName(q)), da_loleakPinData(q)
                        End If
                    End If
                Else
                    Call Print_Error_Message(Warning_Info, VBT_LIB_DC_IDS, "DCVS_IDS_main_auto_range_and_measure", "Wrong Low Leakage information!!!")
                End If
            Else
                Call Print_Error_Message(Warning_Info, VBT_LIB_DC_IDS, "DCVS_IDS_main_auto_range_and_measure", "Wrong Low Leakage information!!!")
            End If
        Next q
    End If
''**** For Special Low Leakage Bin1 Device 20230706 ****

    sOutPut_DisconnectPin = vbNullString
    sOutput_IgnorePin = vbNullString
    sDontCarePin = vbNullString
    sOutput_NCPin = vbNullString
        
    isDoDisconnectPinPass = Disconnect_X_and_Nouse_Pin(DisconnectPinType.UnusedIoPin, patt, EnableDisconnectPins, ExcludeDisconnectPins, sDontCarePin, sOutPut_DisconnectPin, sOutput_IgnorePin, sOutput_NCPin)     'Disconnect UnUse pin after run pattern
    If isDoDisconnectPinPass = False Then Exit Function     'Binout if get something error no need do IDS test
        
    TheHdw.Digital.Patgen.TimeOut = 10
    
    If patt <> "" Then
        'Call TheHdw.Patterns(patt).Load
        '-------------------------------------------DSSC Soruce
        If DigSrc_pin <> "" Then
            rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(patt, patcnt)
            Call GeneralDigSrcSetting(CStr(rtnPatNames(0)), DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, _
                                    DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPwave)
        End If
        '-------------------------------------------
    
        If FlagWait Then
            Call TheHdw.patterns(patt).start
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)  'Meas during CPUA loop
        Else
            Call TheHdw.patterns(patt).test(pfAlways, 0, tlResultModeDomain)
            TheHdw.Digital.Patgen.HaltWait
        End If
    
    Else
        '------------------------------ 191017 Tonga BringUp RTOS Run Scenario Start by Leslie---------------------------------------------------------------------------
        If RTOSCmd <> "" Then
            RTOS_IDS RTOSCmd, RTOSTimeOut
        End If
        '------------------------------ 191017 Tonga BringUp RTOS Run Scenario Start by Leslie---------------------------------------------------------------------------
    End If
    '-------------------------------------------AP & RF
    
    ''[240521]from t-hid
    ' Support Measure after pattern run completed 220802
    If FlagPatFinishMeasure Then
        TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
        TheHdw.Digital.Patgen.HaltWait
    End If
    
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
    End If
        
    Call Disconnect_X_and_Nouse_Pin(DisconnectPinType.Xpin, patt, EnableDisconnectPins, sDontCarePin:=sDontCarePin)   'Disconnect Xpin after run pattern
    
    '-------------------------------------------pre_measure_store
    If Interpose_Meas_before <> "" Then
        Call SetForceCondition(Interpose_Meas_before & ";STOREPREMEAS")
    End If
    '-------------------------------------------
    ''[240521]from t-hid
    'TheHdw.DCVS.Pins("VDD_SRAM_SOC").CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
    
    ''[20240527]auto-range pins split to dcvi dcvs
    Dim pinAry() As String
    Dim pins As Variant
    Dim DCVI_AutoRange_Pin As String
    Dim DCVS_AutoRange_Pin As String
    pinAry() = Split(LCase(AutoRange_Pin), ",")
    For Each pins In pinAry
        
        If LCase(TheExec.DataManager.ChannelType(pins)) Like "*dcvs*" Then
        'If LCase(gl_GetInstrumentType_Dic(Pins)) Like "*dcvs*" Then
            DCVS_AutoRange_Pin = DCVS_AutoRange_Pin & pins & ","
        ElseIf LCase(TheExec.DataManager.ChannelType(pins)) Like "*dcvi*" Then
        'ElseIf LCase(gl_GetInstrumentType_Dic(Pins)) Like "*dcvi*" Then
            DCVI_AutoRange_Pin = DCVI_AutoRange_Pin & pins & ","
        End If
        
    Next pins
    If DCVS_AutoRange_Pin <> "" Then DCVS_AutoRange_Pin = left(DCVS_AutoRange_Pin, Len(DCVS_AutoRange_Pin) - 1)
    If DCVI_AutoRange_Pin <> "" Then DCVI_AutoRange_Pin = left(DCVI_AutoRange_Pin, Len(DCVI_AutoRange_Pin) - 1)
    
    ''[20240530]auto-range pins split to dcvi dcvs
    Set All_Power_data_IDS_GB = Nothing
    If DCVS_Power_Pin <> "" Or DCVS_OtherPower_Pin <> "" Then
        'If (UCase(DCVS_Power_Pin) Like "*CP*=*" Or UCase(DCVS_Power_Pin) Like "*FT*=*") Then DCVS_Power_Pin = Select_MeasPin(DCVS_Power_Pin, UCase(currentJobName))
        'If (UCase(AutoRange_Pin) Like "*CP*=*" Or UCase(AutoRange_Pin) Like "*FT*=*") Then AutoRange_Pin = Select_MeasPin(DCVS_AutoRange_Pin, UCase(currentJobName))
        'If debug_print_pins = vbNullString Then debug_print_pins = DCVS_AutoRange_Pin
        debug_print_pins = DCVS_AutoRange_Pin
        DCVS_IDS_main_auto_range_and_measure DCVS_Power_Pin, DCVS_OtherPower_Pin, All_Power_data, repeat_count, FlowLimitForInitIRange, Search_Step, debug_print_pins, DCVS_AutoRange_Pin, FixCurrentRange, isUse_Product_Identifier, sSpecific_Product_Identifier, Flag, dic_loleakpin            'FixCurrentRange 20220627
    End If
    If DCVI_Power_Pin <> "" Then
        'If (UCase(AutoRange_Pin) Like "*CP*=*" Or UCase(AutoRange_Pin) Like "*FT*=*") Then AutoRange_Pin = Select_MeasPin(DCVI_AutoRange_Pin, UCase(currentJobName))
        'If debug_print_pins = vbNullString Then debug_print_pins = DCVI_AutoRange_Pin
        debug_print_pins = DCVI_AutoRange_Pin
        DCVI_IDS_main_auto_range_and_measure DCVI_Power_Pin, All_Power_data, repeat_count, FlowLimitForInitIRange, Search_Step, debug_print_pins, DCVI_AutoRange_Pin, Flag, dic_loleakpin
    End If

    All_Power_data_IDS_GB = All_Power_data.COPY
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
    
        'Move Patgen.HaltWait before change FRC @William 220809
        
    ''[240521]from t-hid
    ' Support Measure after pattern run completed 220802
    If Not FlagPatFinishMeasure Then
        TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
        TheHdw.Digital.Patgen.HaltWait
    End If
    
    If DisableClock Then
        If DisableFRCPinName <> "" And FRC_RelayPin <> "" Then
            Call Enable_FRC_Pins(DisableFRCPinName, FRC_RelayPin, DisableClock)
        Else
            Call Enable_FRC(DisableClockPortName, DisableClock)
        End If
        'If DebugFlag = True Then TheExec.Datalog.WriteComment "print: nWire connect, pin " & DisableClockPortName
    End If
    

        
        ''20210930 move the IDS calcuation location for validation report
    If DictName <> "" Then Call AddStoredMeasurement(DictName, All_Power_data)
    If Fuse_Enable Then Call IDS_Store2Dic_Mapping(DCVS_Power_Pin, All_Power_data, patt)        'If Fuse_StoreName <> "" Then Call IDS_Store2Dic(Fuse_StoreName, DCVS_Power_Pin, All_Power_data, patt)
    If Calc_Eqn <> "" Then Call ProcessCalcEquation(sTempCalc_Eqn)

    If FlagWait = True Then
        Call HardIP_WriteFuncResult(, , glb_TestInstance)
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
    TheHdw.Digital.ApplyLevelsTiming False, True, False, tlPowered    'SEC DRAM
    '==== Init pin (restore IFold limit) ====

    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, VBT_LIB_DC_IDS, sFuncName)
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function IDS_Delta_calc(Current1_dic As String, Current2_dic As String, power_rail As String)
    ''' Update code for Donan @William 230113
    Dim Current_1 As New PinListData
    Dim Current_2 As New PinListData
    Dim Current_delta As New PinListData
    
    Current_1 = GetStoredMeasurement(Current1_dic)
    Current_2 = GetStoredMeasurement(Current2_dic)
    
    Dim power_rail_ary() As String
    If (UCase(power_rail) Like "*CP*=*" Or UCase(power_rail) Like "*FT*=*") Then power_rail = Select_MeasPin(power_rail, UCase(currentJobName))
    power_rail_ary = Split(power_rail, ",")
    
    Dim p As Variant
'    For Each p In Split(power_rail, ",")
'        Current_delta.AddPin (p)
'        Current_delta.Pins(p).value = Current_1.Pins(p).Subtract(Current_2.Pins(p))
'    Next p
    
    Dim i As Long
    Dim site As Variant
    Dim Tname As String
'    For i = 0 To Current_delta.Pins.Count - 1
'        If InStr(power_rail, Current_delta.Pins.item(i)) <> 0 Then
'            Tname = Report_TName_From_Instance("CalcI", Current_delta.Pins.item(i), Current1_dic & "-" & Current2_dic, , , , , , tlForceNone)
'            TheExec.Flow.TestLimit Current_delta.Pins(i), Tname:=Tname, PinName:=Current_delta.Pins.item(i), ForceResults:=tlForceFlow, lowVal:=0
'        End If
'    Next i
        
    For Each p In power_rail_ary
        Current_delta.AddPin (p)
        Current_delta.pins(p).value = Current_1.pins(p).Subtract(Current_2.pins(p))
        
        Tname = Report_TName_From_Instance("CalcI", p, Current1_dic & "-" & Current2_dic, , , , , , tlForceNone)
        TheExec.flow.TestLimit Current_delta.pins(p), Tname:=Tname, PinName:=p, ForceResults:=tlForceFlow

    Next p
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
    
    Dim pins() As String, Pin_Cnt As Long
    Dim Instance_name_str As String: Instance_name_str = UCase(TheExec.DataManager.instancename)
    Dim Dic_IDS_fuse_name As New Scripting.Dictionary
    Dim site As Variant 'Carter, 20240304
    TheExec.DataManager.DecomposePinList Delta_Pin, pins, Pin_Cnt
    
    If TheExec.CurrentJob = "CP2" Then
     Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_25C
    End If
    
    If TheExec.CurrentJob = "WLFT1" Or TheExec.CurrentJob = "FT1" Then
     Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_85C
    End If
    

    'IDS_from_DCVS = 0 'initial
    If FuseType = "CFG" Then
        
        ''====20201230 add for efuse new code====
        Dim opbank As New eFuseBdfBank
        Dim field As New eFuseBdfField
        Dim fieldStr As Variant
        Dim m_dbl As New SiteDouble
        Dim m_dbl_round As New SiteDouble
        
        Set opbank = GetBdfBank(FuseType)
        
        '20221201, Do pin loop first, fix test limit order
        For j = 0 To UBound(pins())
            '20221215, fields loop TTR, use small ids dic
            For Each field In opbank.DicAllIdsList
            'For Each fieldStr In opbank.Fields
                'Set field = opbank.Fields(fieldStr)
    
                'If TheExec.CurrentJob = "WLFT1" Then
                    If field.Algorithm = alg_ids Then
                        If LCase(field.name) = Dic_IDS_fuse_name(pins(j)) Then
                            IDS_PwrName = pins(j)

                            For Each site In TheExec.sites
                                '20221201, Truncate IDS value with resolution
                                IDS_from_DCVS = All_Power_data_IDS_GB.pins(IDS_PwrName).value
                                
''                               If UCase(TheExec.DataManager.instanceName) Like "IDS_IDSWIFUSERES_PP*" Then
                                        m_dbl = (IDS_from_DCVS.Multiply(1000) / field.Resolution)
                                        m_dbl_round = Round(m_dbl)
                                        If ((CDec(m_dbl_round) - CDec(m_dbl)) >= 0) Then ''''MUST have
                                            m_dbl_round = m_dbl_round
                                        Else
                                            m_dbl_round = m_dbl_round + 1
                                        End If
                                        IDS_from_DCVS = m_dbl_round * field.Resolution * 0.001 '/ A
''                                End If
                            Next site

                            IDS_from_Efuse = field.DsscDecValue.Multiply(field.Resolution * 0.001)
                            ''IDS_from_Efuse = 0.001 ''simulation for CP2, 20230606
                            'IDS_Delta = IDS_from_DCVS.Subtract(IDS_from_Efuse)
                            'IDS_ratio = IDS_from_DCVS.Divide(IDS_from_Efuse)

                             If TheExec.CurrentJob = "CP2" Then
                                IDS_ratio = IDS_from_DCVS.divide(IDS_from_Efuse)
                                TheExec.flow.TestLimit resultVal:=IDS_from_Efuse, Tname:=IDS_PwrName & "_CP1", PinName:=IDS_PwrName & "_CP1", ForceResults:=tlForceFlow
                                TheExec.flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_CP2", PinName:=IDS_PwrName & "_CP2", ForceResults:=tlForceFlow
                                TheExec.flow.TestLimit IDS_ratio, Tname:=IDS_PwrName & "_Ratio", PinName:=IDS_PwrName & "_Ratio", ForceResults:=tlForceFlow
                             End If
                             
                             If TheExec.CurrentJob = "WLFT1" Or TheExec.CurrentJob = "FT1" Then
                                IDS_ratio = IDS_from_Efuse.divide(IDS_from_DCVS)
                                TheExec.flow.TestLimit resultVal:=IDS_from_Efuse, Tname:=IDS_PwrName & "_CP2", PinName:=IDS_PwrName & "_CP2", ForceResults:=tlForceFlow
                                TheExec.flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_WLFT1", PinName:=IDS_PwrName & "_WLFT1", ForceResults:=tlForceFlow
                                TheExec.flow.TestLimit IDS_ratio, Tname:=IDS_PwrName & "_Ratio", PinName:=IDS_PwrName & "_Ratio", ForceResults:=tlForceFlow
                             End If
                        End If
                    End If
                'End If
            Next
        Next j
        
        
    End If
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in DCVS_IDS_main_current_Ratio"
    If AbortTest Then Exit Function Else Resume Next
End Function


''20220124
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
    Dim pins() As String
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
    TheExec.DataManager.DecomposePinList Calc_Pin, pins(), Pin_Cnt
    
    ''[20240527]ids value do Resolution
    Dim m_dbl As New SiteDouble
    Dim m_dbl_round As New SiteDouble
    
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
        ''' Update for Donan to fix incorrect IDS stage  @william 230131
        For j = 0 To UBound(IDS_MAPPING)
            If UCase(FusedStage) = UCase(IDS_MAPPING(j).stage) Then
                CurrentJob_IDSInfo = IDS_MAPPING(j)
                Exit For
            End If
        Next j
        
        Set opbank = GetBdfBank(FuseType)
        For j = 0 To UBound(pins)
            ''link pin name with field name.
            IDS_PwrName = CurrentJob_IDSInfo.MappingDict.item(pins(j))
            Set field = opbank.Fields(IDS_PwrName)
            
            ''get measured value and fused value
            IDS_from_DCVS = All_Power_data_IDS_GB.pins(pins(j))
            
            ''[20240527]ids value do Resolution
            For Each site In TheExec.sites
                '20221201, Truncate IDS value with resolution
'                IDS_from_DCVS = All_Power_data_IDS_GB.Pins(IDS_PwrName).value
                
''              If UCase(TheExec.DataManager.instanceName) Like "IDS_IDSWIFUSERES_PP*" Then
                    m_dbl = (IDS_from_DCVS.Multiply(1000) / field.Resolution)
                    m_dbl_round = Round(m_dbl)
                    If ((CDec(m_dbl_round) - CDec(m_dbl)) >= 0) Then ''''MUST have
                        m_dbl_round = m_dbl_round
                    Else
                        m_dbl_round = m_dbl_round + 1
                    End If
                    IDS_from_DCVS = m_dbl_round * field.Resolution * 0.001 '/ A
''              End If
            Next site
            
            If SumField.Exists(pins(j)) Then
                'get sum fused ids value
                '20240309 - fix get wrong data with GetStoredMeasurement, use GetStoredData
                 IDS_from_Efuse = GetStoredData(pins(j) + "_EFUSE")
'                IDS_from_Efuse = GetStoredMeasurement(Pins(j) + "_EFUSE")
            Else
                If TheExec.TesterMode = testModeOffline Then
                    IDS_from_Efuse = IDS_Math.Add(0.01 + Rnd() * 0.0001)
                Else
                    IDS_from_Efuse = field.DsscDecValue.Multiply(field.Resolution * 0.001)  '20210406 Modify for new Efuse
                End If
            End If
            
            
            ''' Remove read value low limit @William 231116
'            TheExec.Flow.TestLimit resultVal:=IDS_from_Efuse, lowVal:=0.0001, Tname:=IDS_PwrName & "_" & FusedStage, PinName:=Pins(j) & "_" & FusedStage, ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=IDS_from_Efuse, Tname:=IDS_PwrName & "_" & FusedStage, PinName:=pins(j) & "_" & FusedStage, scaletype:=scaleMilli, unit:=unitAmp, formatStr:="%.3f", ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_" & currentJobName, PinName:=pins(j) & "_" & currentJobName, scaletype:=scaleMilli, unit:=unitAmp, formatStr:="%.3f", ForceResults:=tlForceNone
            ''' Update for Hidra. Set 'mA' for IDS unit @William 240220
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment _
                "site(" & site & ") FusedStage = " & FusedStage & ", eFuseFieldName = " & IDS_PwrName & ", FusedValue = " & Format(IDS_from_Efuse * 1000, "0.000000") & "mA" & _
                                ", MeasuredStage = " & currentJobName & ", PinName = " & pins(j) & ", MeasuredValue = " & Format(IDS_from_DCVS * 1000, "0.000000") & "mA"
            Next

            Select Case (MathFunction)
                Case "+":
                    IDS_Math = IDS_from_DCVS.Add(IDS_from_Efuse)
                    TheExec.flow.TestLimit IDS_Math, Tname:=IDS_PwrName & "_Add", PinName:=pins(j) & "_Add", ForceResults:=tlForceFlow
                Case "-":
                    IDS_Math = IIf(Not MathInverse, IDS_from_DCVS.Subtract(IDS_from_Efuse), IDS_from_Efuse.Subtract(IDS_from_DCVS))
                    TheExec.flow.TestLimit IDS_Math, Tname:=IDS_PwrName & "_Delta", PinName:=pins(j) & "_Delta", ForceResults:=tlForceFlow
                Case "*":
                    IDS_Math = IIf(Not MathInverse, IDS_from_DCVS.Multiply(IDS_from_Efuse), IDS_from_Efuse.Multiply(IDS_from_DCVS))
                    TheExec.flow.TestLimit IDS_Math, Tname:=IDS_PwrName & "_Mul", PinName:=pins(j) & "_Mul", ForceResults:=tlForceFlow, unit:=unitNone
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
                    TheExec.flow.TestLimit IDS_Math, Tname:=IDS_PwrName & "_Ratio", PinName:=pins(j) & "_Ratio", ForceResults:=tlForceFlow, unit:=unitNone
                    For Each site In TheExec.sites
                        If FlagChk = False Then
                            TheExec.Datalog.WriteComment "ERROR!! Site(" & site & ") the denominator of the ratio is 0."
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

Public Function Parse_IDS_Mapping_Table() As Long
On Error GoTo errHandler
    
    Call ParseIDSMappingTable(True)
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_IDS", "Parse_IDS_Mapping_Table") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Check_IDS_Result_With_BinCut(IDS_Pins As PinList, FuseType As String, CheckJobType As String)   'check IDS result and bincut value
    On Error GoTo errHandler
        Dim opbank As New eFuseBdfBank
        Dim Dic_IDS_fuse_name As Dictionary
        Dim pins() As String
        Dim field As New eFuseBdfField
        Dim IDS_from_Efuse As New SiteDouble
        Dim pin As Variant
        Dim Hilimit_IDS As New SiteDouble
        Dim powerPin As String
        Dim highest_Pmode_Name As String
        
        If Not UCase(FuseType) = "CFG" Then Exit Function
        
        'Call ReadEfuseDataFromBinCut   'test use
        
        Set opbank = GetBdfBank(FuseType)
        
        '01.Get IDS efuse Dictionary
        If CheckJobType = "CP2" Then
            Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_85C    'CP2 efuse store place
        Else
            Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_25C    'CP1 efuse store place
        End If
        
        pins = Split(IDS_Pins, ",")
        
        For Each pin In pins
            '02.Read efuse IDS value
            Set field = opbank.Fields(LCase(Dic_IDS_fuse_name(UCase(pin))))
            IDS_from_Efuse = field.DsscDecValue.Multiply(field.Resolution * 0.001)

            Hilimit_IDS = 0#

            If domain2pinDict.Exists(UCase(pin)) Then
                powerPin = UCase(pin)
            ElseIf pin2domainDict.Exists(UCase(pin)) Then
                powerPin = VddbinPin2Domain(UCase(pin))
            Else
                If UCase(pin) = "VDD_AMPH_DDR" Then
                    powerPin = vbNullString
                Else
                    TheExec.Datalog.WriteComment pin & " is not BinCut CorePower or OtherRail. Please check Check_IDS_Result_With_BinCut.Error!!!"
                    TheExec.ErrorLogMessage pin & " is not BinCut CorePower or OtherRail. Please check Check_IDS_Result_With_BinCut.Error!!!"
                    TheExec.flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:=pin & " is not BinCut CorePower or OtherRail. Please check Check_IDS_Result_With_BinCut.Error!!!"
                End If
            End If

            If UCase("*," & FullCorePowerinFlowSheet & ",*") Like UCase("*," & powerPin & ",*") Then '''corePower

                If dict_IsCorePowerInBinCutFlowSheet.item(powerPin) = True And UCase(powerPin) Like "*VDD*_SRAM*" Then
                    highest_Pmode_Name = BinCut_Sram_Power_Seq("MPS")(UBound(BinCut_Sram_Power_Seq("MPS")))
                Else
                    highest_Pmode_Name = BinCut_Power_Seq(VddBinStr2Enum(powerPin)).Power_Seq(UBound(BinCut_Power_Seq(VddBinStr2Enum(powerPin)).Power_Seq))
                End If

                '03.get the bincut limit value by BIN1
                For Each site In TheExec.sites
                    If CurrentPassBinCutNum_normal(site) = 1 Then
                        If CheckJobType = "CP2" Then
                            Hilimit_IDS(site) = BinCut(VddBinStr2Enum(highest_Pmode_Name), 1).IDS_FT_LIMIT(0) / 1000
                        Else
                            Hilimit_IDS(site) = BinCut(VddBinStr2Enum(highest_Pmode_Name), 1).IDS_CP_LIMIT(0) / 1000
                        End If
                        'check if IDS larger then bincut value get fail
                        TheExec.flow.TestLimit resultVal:=IDS_from_Efuse(site), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=CStr(pin) & "_" & CheckJobType, PinName:=CStr(pin) & "_" & CheckJobType, hiVal:=Hilimit_IDS(site)
                    Else
                        'if not BIN1 no need check
                        TheExec.flow.TestLimit resultVal:=IDS_from_Efuse(site), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=CStr(pin) & "_" & CheckJobType, PinName:=CStr(pin) & "_" & CheckJobType
                    End If
                Next site
                
            ElseIf UCase("*," & FullOtherRailinFlowSheet & ",*") Like UCase("*," & powerPin & ",*") Then  '''non corePower
                '03.get the bincut limit value by BIN1
                For Each site In TheExec.sites
                    If CurrentPassBinCutNum_normal(site) = 1 Then
                        If CheckJobType = "CP2" Then
                            Hilimit_IDS(site) = FTIDS_Spec(VddBinStr2Enum(powerPin), 1) / 1000 '220629 CP2 IDS change to Bin1 spec
                        Else
                            Hilimit_IDS(site) = CPIDS_Spec(VddBinStr2Enum(powerPin), 1) / 1000
                        End If
                        'check if IDS larger then bincut value get fail
                        TheExec.flow.TestLimit resultVal:=IDS_from_Efuse(site), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=CStr(pin) & "_" & CheckJobType, PinName:=CStr(pin) & "_" & CheckJobType, hiVal:=Hilimit_IDS(site)
                    Else
                        'if not BIN1 no need check
                        TheExec.flow.TestLimit resultVal:=IDS_from_Efuse(site), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=CStr(pin) & "_" & CheckJobType, PinName:=CStr(pin) & "_" & CheckJobType
                    End If
                Next site
            Else
                If UCase(pin) = "VDD_AMPH_DDR" Then
                    Hilimit_IDS = 0
                Else
                    TheExec.Datalog.WriteComment pin & " is not BinCut CorePower or OtherRail. Please check Check_IDS_Result_With_BinCut.Error!!!"
                    TheExec.ErrorLogMessage pin & " is not BinCut CorePower or OtherRail. Please check Check_IDS_Result_With_BinCut.Error!!!"
                    TheExec.flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:=pin & " is not BinCut CorePower or OtherRail. Please check Check_IDS_Result_With_BinCut.Error!!!"
                End If
            End If

        Next pin
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function IDS_main_current_Binning_new(MathPins As String, Spec As String, Optional Flag As String, Optional isCompareEfuseValue As Boolean = False, Optional FuseType As String = vbNullString, Optional SpecificBinCutNum As Long = 1)

    On Error GoTo errHandler
    
    Dim i, j As Long
    Dim TempPin As String
    Dim Pin_Cnt As Long
    Dim pins() As String
    Dim Spec_Ary() As String
    Dim Spec_Ary_Val() As String
    Dim Flag_ary() As String
    Dim IDS_from_DCVS As New SiteDouble
    Dim Hilimit_IDS As Double
    Dim opbank As New eFuseBdfBank
    Dim field As New eFuseBdfField
    Dim CurrentJob_IDSInfo As IDS_Mapping_Info
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim sWarningMsg As String
    Dim dic_Spec As New Dictionary
    Dim Val_Hi() As String

    Call TheExec.flow.GetTestLimits(FlowLimitsInfo)

    
    If FlowLimitsInfo Is Nothing Then
    Else
        FlowLimitsInfo.GetHighLimits Val_Hi
    End If
    
    TempPin = Replace(MathPins, "+", ",")
    TheExec.DataManager.DecomposePinList TempPin, pins(), Pin_Cnt
    
    Spec_Ary = Split(Spec, "+")
    Flag_ary = Split(Flag, "+")

    ReDim Preserve Val_Hi(UBound(pins))

    'Set Pin and Spec dictionary
    If UBound(Spec_Ary) >= 0 And UBound(pins) = UBound(Spec_Ary) Then
        For i = 0 To UBound(Spec_Ary)
            Spec_Ary_Val = Split(Spec_Ary(i), ":")
            If UBound(Spec_Ary_Val) = 1 Then
                If dic_Spec.Exists(Spec_Ary_Val(0)) Then
                Else
                    dic_Spec.Add Spec_Ary_Val(0), Spec_Ary_Val(1)
                End If
            End If
        Next i
    ElseIf UBound(Spec_Ary) = 0 Then
    Else
        sWarningMsg = "Pin and Spec count not match!!"
        Call Print_Error_Message(Warning_Info, VBT_LIB_DC_IDS, "IDS_main_current_Binning", sWarningMsg)
        Exit Function
    End If

    'get same stage IDS mapping table name array
    If isCompareEfuseValue = True And FuseType <> "" Then
        Set opbank = GetBdfBank(FuseType)

        For i = 0 To UBound(IDS_MAPPING)
            If UCase(IDS_MAPPING(i).stage) = "CP1" Then     'just get CP1 IDS power pin in efuse name
                CurrentJob_IDSInfo = IDS_MAPPING(i)
                Exit For
            End If
        Next i
    End If

    'check bincut sheet exist, if not use will return 999
    SpecificBinCutNum = CheckHaveBinCutSheet(SpecificBinCutNum, False)

    For i = 0 To UBound(pins)
        If gl_dicPowerPinIndex.Exists(LCase(pins(i))) Then
            If isCompareEfuseValue = True Then
                Set field = opbank.Fields(CurrentJob_IDSInfo.MappingDict(LCase(pins(i))))   'use efuse pin name to get IDS result in efuse
                IDS_from_DCVS = field.MeasureValue.divide(1000)
            Else
                'Get Measurement Value from IDS
                IDS_from_DCVS = gl_All_Power_data_IDS_CP1.pins(pins(i))    'use current test result
            End If
            If dic_Spec.Exists(pins(i)) Then
                Hilimit_IDS = Compare_ForceVal_BV(pins(i), Val_Hi(i), nBincutNum:=SpecificBinCutNum)          'get BinCut sheet limit if use limit is ""
                Hilimit_IDS = Hilimit_IDS * Evaluate(dic_Spec(pins(i)))       'BinCut sheet limit * ratio
            Else    'if power pin spec not exist that ratio = 1
                Hilimit_IDS = Compare_ForceVal_BV(pins(i), Val_Hi(i), nBincutNum:=SpecificBinCutNum)       'get BinCut sheet limit if use limit is ""
            End If

            TheExec.flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=pins(i) & "_" & "Judge_IDS_Current_Spec", PinName:=pins(i) & "_" & currentJobName, hiVal:=Hilimit_IDS, unit:=a, ForceResults:=tlForceNone
            'Clear Variant
            Set IDS_from_DCVS = Nothing
        Else
            sWarningMsg = "Pin = " & pins(i) & " is N/C pin."
            Call Print_Error_Message(Warning_Info, VBT_LIB_DC_IDS, "IDS_main_current_Binning", sWarningMsg)
        End If
    Next i

    'Printing Flag state if argument is not ""
    If Flag <> "" Then
        For i = 0 To UBound(Flag_ary)
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "Site[" & site & "]" & Flag_ary(i) & "=" & TheExec.sites.item(site).FlagState(Flag_ary(i))
            Next site
        Next i
    End If

    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, VBT_LIB_DC_IDS, "IDS_main_current_Binning")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IDS_main_current_Binning(MathPins As String, Spec As String, Optional Flag As String)

    Dim i, j As Long
    On Error GoTo errHandler
    Dim pins() As String
    Dim Spec_Ary() As String
    Dim Spec_Ary_Val() As String
    Dim Flag_ary() As String
    Dim IDS_from_DCVS As New SiteDouble
    Dim Hilimit_IDS As Double
    Dim opbank As New eFuseBdfBank
    Dim field As New eFuseBdfField
    Dim Dic_IDS_fuse_name As Dictionary

    pins = Split(MathPins, "+")
    Spec_Ary = Split(Spec, "+")
    Flag_ary = Split(Flag, "+")
    Set opbank = GetBdfBank("CFG")
    
    If UCase(TheExec.CurrentJob) = "CP1" Then
        Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_25C
    ElseIf UCase(TheExec.CurrentJob) = "CP2" Then
        Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_85C
    End If

    For i = 0 To UBound(pins)
        'Get Measurement Value from IDS
        Set field = opbank.Fields(Dic_IDS_fuse_name(UCase(pins(i))))
        IDS_from_DCVS = field.MeasureValue.divide(1000)   'All_Power_data_IDS_GB_CORR.Pins(Pins(i))
        
        
    If UCase(TheExec.CurrentJob) = "CP1" Then
        'Get IDS BinCut Hilimit from global arrary
        For j = 0 To UBound(ids_range_ary_GB)
            If pins(i) = ids_range_ary_GB(j).PinName Then
                Hilimit_IDS = ids_range_ary_GB(j).hiLimit
                Exit For
            End If
        Next j
        
        Spec_Ary_Val = Split(Spec_Ary(i), ":")     'VDD_PCPU:5/6
        If Spec_Ary_Val(0) = pins(i) Then
            Hilimit_IDS = Hilimit_IDS * Evaluate(Spec_Ary_Val(1))
        End If
     End If
     
        TheExec.flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=pins(i) & "_" & "Judge_IDS_Current_Spec", PinName:=pins(i) & "_" & currentJobName, hiVal:=Hilimit_IDS, ForceResults:=tlForceFlow
        'Clear Variant
        Set IDS_from_DCVS = Nothing
    Next i
    'Printing Flag state if argument is not ""
    If Flag <> "" Then
        For i = 0 To UBound(Flag_ary)
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "Site[" & site & "]" & Flag_ary(i) & "=" & TheExec.sites.item(site).FlagState(Flag_ary(i))
            Next site
        Next i
    End If
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in IDS_main_current_Binning"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Clear_IDS_Dic_for_each_die()

gl_IDS_INFO_Dic.RemoveAll
gl_IDS_INFO_AllCore_Dic.RemoveAll
gl_IDS_INFO_4Core_Dic.RemoveAll

End Function

Public Function DCVS_IDS_main_current_Read(IDS_Pins As PinList, FuseType As String)
On Error GoTo errHandler
    Dim pin As Variant
    Dim opbank As New eFuseBdfBank
    Dim Dic_IDS_fuse_name As Dictionary
    Dim IDS_from_Efuse As New SiteDouble
    Dim field As New eFuseBdfField
    Dim pins() As String
    
    If Not UCase(FuseType) = "CFG" Then Exit Function
    
    Set opbank = GetBdfBank(FuseType)
    Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_25C
    pins = Split(IDS_Pins, ",")
    
    For Each pin In pins
        Set field = opbank.Fields(Dic_IDS_fuse_name(UCase(pin)))
        IDS_from_Efuse = field.DsscDecValue.Multiply(field.Resolution * 0.001)
        TheExec.flow.TestLimit resultVal:=IDS_from_Efuse, Tname:=CStr(pin) & "_CP1", PinName:=CStr(pin) & "_CP1", ForceResults:=tlForceFlow
    Next pin
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in DCVS_IDS_main_current_Read"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DCVS_IDS_main_current_Read_CP2(IDS_Pins As PinList, FuseType As String)
On Error GoTo errHandler
    Dim pin As Variant
    Dim opbank As New eFuseBdfBank
    Dim Dic_IDS_fuse_name As Dictionary
    Dim IDS_from_Efuse As New SiteDouble
    Dim field As New eFuseBdfField
    Dim pins() As String
    
    If Not UCase(FuseType) = "CFG" Then Exit Function
    
    Set opbank = GetBdfBank(FuseType)
    Set Dic_IDS_fuse_name = gl_Dic_IDS_fuse_name_AllCore_85C
    pins = Split(IDS_Pins, ",")
    
    For Each pin In pins
        Set field = opbank.Fields(Dic_IDS_fuse_name(UCase(pin)))
        IDS_from_Efuse = field.DsscDecValue.Multiply(field.Resolution * 0.001)
        'TheExec.Flow.TestLimit resultVal:=IDS_from_Efuse, Tname:=CStr(pin) & "_CP2", PinName:=CStr(pin) & "_CP2", ForceResults:=tlForceFlow
        TheExec.flow.TestLimit resultVal:=IDS_from_Efuse, Tname:="Check_ids_" & CStr(pin) & "_CP2", ForceResults:=tlForceFlow
    Next pin
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in DCVS_IDS_main_current_Read"
    If AbortTest Then Exit Function Else Resume Next
End Function



