Attribute VB_Name = "LIB_HardIP_ForceCondition"
#Const isUFP = True
Option Explicit

Public CHAR_USL_HVCC As Double
Public CHAR_USL_LVCC As Double
Public CHAR_LSL_HVCC As Double
Public CHAR_LSL_LVCC As Double
Public Charz_Power_condition  As String
Public Charz_Force_Power_condition  As String
Public CharSetName_GLB As String
Public Re_store As Double
Public g_PPMU_Connected As String
Public PrePat_Restore_String As String
Public PreMeas_Restore_String As String
Public PrePatStore As Boolean
Public PreMeasStore As Boolean

'================================================================================
'180425 update for trace compensation
Public gTerm_cond_All As String
Public gTerm_Restore_cond As String
Public gTerm_cond_Flag As Boolean
Public gTerm_cond_shm_Flag As Boolean


Public f_restorePrePat As Boolean

'[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
Dim SD_Shmoo_GLB_Val As New SiteDouble
'================================================================================
'Public Function Force_Condition_V(force_pin As String, force_val As Double, Optional restorepin As Boolean)
'On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'    Dim p_ary() As String, p_cnt As Long
'    p_ary = Split(force_pin, ",") ' added to allow pinlist with comma such as "VDD_CPU,VDD_SRAM"
'    If (LCase(TheExec.DataManager.PinType(LCase(p_ary(0)))) Like "power" Or LCase(TheExec.DataManager.PinType(LCase(p_ary(0)))) Like "analog") Then
'        SetPowerValue force_pin, force_val
'        If restorepin = True And LCase(TheExec.DataManager.PinType(LCase(p_ary(0)))) Like "analog" = True Then
'            With TheHdw.DCVI.Pins(force_pin)
'                .Gate = False
'                .Disconnect
'            End With
'        End If
'    Else
'        TheHdw.Digital.Pins(force_pin).Disconnect
'        If (PrePatStore = True) Then
'            If (PrePat_Restore_String = "") Then
'                PrePat_Restore_String = force_pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Voltage, "0.000000"))
'            Else
'                PrePat_Restore_String = PrePat_Restore_String & ";" & force_pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Voltage, "0.000000"))
'            End If
'        ElseIf (PreMeasStore = True) Then
'            If (PreMeas_Restore_String = "") Then
'                PreMeas_Restore_String = force_pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Voltage, "0.000000"))
'            Else
'                PreMeas_Restore_String = PreMeas_Restore_String & ";" & force_pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Voltage, "0.000000"))
'            End If
'        Else
'        'Do nothing
'        End If
'
'        If (PrePatStore = True) Then
'            If (PrePat_Restore_String = "") Then
'                PrePat_Restore_String = force_pin + ":DisConnectPPMU;" + force_pin + ":ConnectDigital"
'            Else
'                PrePat_Restore_String = PrePat_Restore_String + ";" + force_pin + ":DisConnectPPMU;" + force_pin + ":ConnectDigital"
'            End If
'        ElseIf (PreMeasStore = True) Then
'            If (PreMeas_Restore_String = "") Then
'                PreMeas_Restore_String = force_pin + ":DisConnectPPMU;" + force_pin + ":ConnectDigital"
'            Else
'                PreMeas_Restore_String = PreMeas_Restore_String + ";" + force_pin + ":DisConnectPPMU;" + force_pin + ":ConnectDigital"
'            End If
'        Else
'        'Do nothing
'        End If
'
'        With TheHdw.PPMU.Pins(force_pin)
'            If force_val = -999 Then
'                '[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
'                For Each site In TheExec.sites
'                    .ForceV SD_Shmoo_GLB_Val, 0.02
'                Next site
'            Else
'                .ForceV (force_val), 0.02
'            End If
'            .Connect
'            .Gate = tlOn
'        End With
'
'        If g_PPMU_Connected <> "" Then
'            g_PPMU_Connected = g_PPMU_Connected & "," & force_pin
'        Else
'            g_PPMU_Connected = force_pin
'        End If
'    End If
'Exit Function 'Add ErrHandler 2023/05/29
'errHandler: 'Add ErrHandler 2023/05/29
'    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Force_Condition_V") 'Add ErrHandler 2023/05/29
'    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
'End Function
'20170419 add for case I
'Public Function Force_Condition_I(force_pin As String, force_val As Double)
'On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'    If (LCase(TheExec.DataManager.PinType(LCase(force_pin))) Like "power") Or (LCase(TheExec.DataManager.PinType(LCase(force_pin))) Like "analog") Then
'        SetPowerValue_I force_pin, force_val
'    Else
'        TheHdw.Digital.Pins(force_pin).Disconnect
'        If (PrePatStore = True) Then
'            If (PrePat_Restore_String = "") Then
'                PrePat_Restore_String = force_pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Current, "0.000000"))
'            Else
'                PrePat_Restore_String = PrePat_Restore_String & ";" & force_pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Current, "0.000000"))
'            End If
'        ElseIf (PreMeasStore = True) Then
'            If (PreMeas_Restore_String = "") Then
'                PreMeas_Restore_String = force_pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Current, "0.000000"))
'            Else
'                PreMeas_Restore_String = PreMeas_Restore_String & ";" & force_pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Current, "0.000000"))
'            End If
'        Else
'        'Do nothing
'        End If
'
'        If (PrePatStore = True) Then
'            If (PrePat_Restore_String = "") Then
'                PrePat_Restore_String = force_pin + ":DisConnectPPMU"
'            Else
'                PrePat_Restore_String = PrePat_Restore_String + ";" + force_pin + ":DisConnectPPMU"
'            End If
'        ElseIf (PreMeasStore = True) Then
'            If (PreMeas_Restore_String = "") Then
'                PreMeas_Restore_String = force_pin + ":DisConnectPPMU"
'            Else
'                PreMeas_Restore_String = PreMeas_Restore_String + ";" + force_pin + ":DisConnectPPMU"
'            End If
'        Else
'        'Do nothing
'        End If
'
'        With TheHdw.PPMU.Pins(force_pin)
'            '.mode = tlPPMUForceIMeasureV
'            .ForceI (force_val)
'            .Connect
'            .Gate = tlOn
'        End With
'
'        If g_PPMU_Connected <> "" Then
'            g_PPMU_Connected = g_PPMU_Connected & "," & force_pin
'        Else
'            g_PPMU_Connected = force_pin
'        End If
'    End If
'Exit Function 'Add ErrHandler 2023/05/29
'errHandler: 'Add ErrHandler 2023/05/29
'    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Force_Condition_I") 'Add ErrHandler 2023/05/29
'    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
'End Function


Public Function SetForceCondition(Setup_string As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim Cat_temp() As String
Dim cat_split_temp As Variant
Dim pin_restore_analog As Boolean
Dim Pin_info_temp() As String

Dim Edge_and_TimeSet() As String    '20180702 add

Dim PinName As String, Pin_Type As String, Pin_value As String, pin_restore As Boolean
Dim Split_cnt As Integer
Dim ins_name As String
Dim site As Variant


g_Retention_ForceV = vbNullString
g_Retention_VDD = vbNullString
'CharSetName_GLB = "" ''avoid VBT error if HIP universal(Meas_FreqVoltCurr_Universal_func) wants to do the shmoo by opcode "test"
'============================customer===================================================================
g_PPMU_Connected = vbNullString

'===============================================================================================
    'get instance name
    ins_name = TheExec.DataManager.instancename
'===============================================================================================
If InStr(Setup_string, "[") > 0 Then ''e.g. PinA:V:0.5*[SrcCodeIndx]+0.001*[SrcCodeIndx1]
        Dim FlowVarStr As String
        Dim TempIdx As Long
        Dim SplitLeftBigColon() As String
        SplitLeftBigColon = Split(Setup_string, "[")
        For TempIdx = 0 To UBound(SplitLeftBigColon)
            If InStr(SplitLeftBigColon(TempIdx), "]") > 0 Then
                FlowVarStr = Split(SplitLeftBigColon(TempIdx), "]")(0)
                Setup_string = Replace(Setup_string, "[" & FlowVarStr & "]", CStr(TheExec.Flow.var(FlowVarStr).value))
            End If
        Next TempIdx
End If

    '==== Use for Sweep Voltage function ==== -- 20221018
    'Ex: CPU_0_MON22:V:sweepvoltage ==> CPU_0_MON22:V:0.1
    If InStr(LCase(Setup_string), "sweepvoltage") <> 0 Then
        TheExec.Datalog.WriteComment "[SweepVoltage]: " & Setup_string
        Setup_string = Replace(Setup_string, "sweepvoltage", gl_sweepVoltage_Value)
    End If

'180425 update for trace compensation
    Setup_string = LCase(Replace(Setup_string, " ", vbNullString))
'    Analyze_shmoo_setup ' for trace compensation
    Call Add_Term_Restore(Setup_string)  ' for trace compensation
    If (Setup_string = "restorepremeas") Then
        Setup_string = LCase(PreMeas_Restore_String)
        PreMeas_Restore_String = vbNullString
        If (Setup_string <> "") Then TheExec.Datalog.WriteComment "restore premeas Force Condtion:" & Setup_string
        pin_restore_analog = True
    ElseIf (Setup_string = "restoreprepat_term") Then
        Setup_string = LCase(PrePat_Restore_String) & ":term" 'remain keyword "term" in order to impact Add_Term_Resotore function
        PrePat_Restore_String = vbNullString
        If (Setup_string <> "") Then TheExec.Datalog.WriteComment "restore prepat Force Condtion:" & Setup_string
        pin_restore_analog = True
    ElseIf (Setup_string = "restoreprepat") Then
        Setup_string = LCase(PrePat_Restore_String)
        PrePat_Restore_String = vbNullString
        If (Setup_string <> "") Then TheExec.Datalog.WriteComment "restore prepat Force Condtion:" & Setup_string
        pin_restore_analog = True
    Else
    'Do nothing
    End If
    
    
    If (UCase(Setup_string) Like "*STOREPREMEAS") Then
        PreMeasStore = True
        PreMeas_Restore_String = vbNullString
        Setup_string = Replace(UCase(Setup_string), ";STOREPREMEAS", vbNullString)
        If TheExec.DevChar.Setups.IsRunning = True Then ' 20180702 add
            Charz_Force_Power_condition = Setup_string
        End If
    ElseIf (UCase(Setup_string) Like "*STOREPREPAT") Then
        PrePatStore = True
        PrePat_Restore_String = vbNullString
        Setup_string = Replace(UCase(Setup_string), ";STOREPREPAT", vbNullString)
        If TheExec.DevChar.Setups.IsRunning = True Then ' 20180702 add
            Charz_Force_Power_condition = Setup_string
        End If
    Else
    'Do nothing
    End If
'===============================================================================================
    '==== Pin Norestore function for DCVS and DCVI ====
    Dim NoRestore_Flag As Long: NoRestore_Flag = 0  '0:No Change(Default), 1:PrePat PinNoRestore, 2:PreMeas PinNoRestore

    If (Setup_string = "") Then Exit Function

    Cat_temp = Split(Setup_string, ";")    'compatible with Autogen and used in central
    
    
    TheExec.Datalog.WriteComment "Force Condtion:" & Setup_string
    Dim flag_shmoo_set_current_point As Boolean
    flag_shmoo_set_current_point = True
    For Each cat_split_temp In Cat_temp
    
        If cat_split_temp = "" Then GoTo continue1
        
        pin_restore = False
        
        If (InStr(LCase(cat_split_temp), "restore") > 0) Then
            pin_restore = True
        End If
        
        Pin_info_temp = Split(cat_split_temp, ":")
        Split_cnt = UBound(Pin_info_temp) + 1
        
        '20230714 sync RF feature
        #If RF = True Then
                If (InStr(UCase(cat_split_temp), "RUNINITPAT") > 0) Then
                        HIPUtility.Initialize inStr_DsscSetup:=glb_DSSCSetup
                        Dim argv As String
                        Dim s_temp_cat As String
                        s_temp_cat = UCase(cat_split_temp)
                        argv = Replace(s_temp_cat, "RUNINITPAT(", "")
                        argv = Replace(argv, ")", "")
                        HIPUtility.RunInitPat_testresult (argv)
                        GoTo continue1
                End If
        #End If
        
        '\\\\\\\\\\\\\\\SAVE H/L Limit\\\\\\\\\\\\\
        If ((LCase(cat_split_temp) Like "usl*") Or LCase(cat_split_temp) Like "lsl*") Then
            If (Split_cnt = 1) Then
                GoTo continue1
            ElseIf (Split_cnt = 2) Then
                Pin_info_temp(1) = CStr(Spec_Evaluate_DC_for_flow_loop(Pin_info_temp(0), "V", Pin_info_temp(1)))
                '\\\\\\Save HVCC Limit\\\\\\
                If (Pin_info_temp(0) = "USL") Then
                    If (Pin_info_temp(1) = "") Then
                        CHAR_USL_HVCC = 9999
                    Else
                        CHAR_USL_HVCC = FormatNumber(CDbl(Pin_info_temp(1)), 3)
                        CHAR_USL_LVCC = FormatNumber(CDbl(Pin_info_temp(1)), 3)
                    End If
                    GoTo continue1
                End If
                '\\\\\\Save LVCC Limit\\\\\\
                If (Pin_info_temp(0) = "LSL") Then
                    If (Pin_info_temp(1) = "") Then
                        CHAR_LSL_LVCC = 9999
                    Else
                        CHAR_LSL_HVCC = FormatNumber(CDbl(Pin_info_temp(1)), 3)
                        CHAR_LSL_LVCC = FormatNumber(CDbl(Pin_info_temp(1)), 3)
                    End If
                End If
            Else
            'Do nothing
            End If
            GoTo continue1
        End If
        
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\Read CharSetName_GLB\\\\\\
        If (InStr(LCase(cat_split_temp), "charsetname") > 0) Then
            CharSetName_GLB = Pin_info_temp(1)
            GoTo continue1
        End If
        
        Dim InStrTmp As String

        
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\20161229 Roy Modified for  Evaluate\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If (UBound(Pin_info_temp) >= 2) Then
            PinName = Pin_info_temp(0)
            Pin_Type = Pin_info_temp(1)
            Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(Pin_info_temp(0), Pin_info_temp(1), Pin_info_temp(2)))
        ElseIf (UBound(Pin_info_temp) = 1) Then
            PinName = Pin_info_temp(0)
            Pin_Type = vbNullString
            Pin_value = CStr(Pin_info_temp(1))
        Else
        'Do nothing
        End If
        
        '////////////////////////////////////////////////////////////////////////////////////////////////
        '==== Pin Norestore function for DCVS and DCVI ==== '20221026
        If (UBound(Pin_info_temp) >= 3) Then
            If InStr(UCase(Pin_info_temp(3)), "NORESTORE") <> 0 And PrePatStore = True Then NoRestore_Flag = 1     'PrePat_NoRestore
            If InStr(UCase(Pin_info_temp(3)), "NORESTORE") <> 0 And PreMeasStore = True Then NoRestore_Flag = 2    'PreMeas_NoRestore
        Else
            NoRestore_Flag = 0
        End If
        '///////////////////////////////////////// Sweep TName //////////////////////////////////////////
        
        If LCase(Pin_info_temp(0)) = "sweep_name" Then
            If gl_UseStandardTestName_Flag = True Then
                gl_Sweep_Name = Pin_info_temp(1)
                If (InStr(LCase(PrePat_Restore_String), "clear_gl_tname") = 0) Then
                    If (PrePat_Restore_String = "") Then
                        PrePat_Restore_String = "gl_tname:clear"
                    Else
                        PrePat_Restore_String = PrePat_Restore_String & ";" & "gl_tname:clear"
                    End If
                End If
            End If
            GoTo continue1
        End If
        
        If LCase(Pin_info_temp(0)) = "sweepy_name" Then
            If gl_UseStandardTestName_Flag = True Then
                gl_SweepY_Name = Pin_info_temp(1)
                If (InStr(LCase(PrePat_Restore_String), "clear_gl_tname") = 0) Then
                    If (PrePat_Restore_String = "") Then
                        PrePat_Restore_String = "gl_tname:clear"
                    Else
                        PrePat_Restore_String = PrePat_Restore_String & ";" & "gl_tname:clear"
                    End If
                End If
            End If
            GoTo continue1
        End If
        
        If LCase(Pin_info_temp(0)) = "gl_tname" And Pin_info_temp(1) = "clear" Then
            If gl_UseStandardTestName_Flag = True Then
                gl_SweepY_Name = vbNullString
                gl_Tname_Meas = vbNullString
                gl_Tname_Alg = vbNullString
                gl_Sweep_Name = vbNullString
                gl_Tname_Alg_Index = 0
            End If
            GoTo continue1
        End If
        
        '////////////////////////////////////////////////////////////////////////////
        '20191217 ChrisHsu modify for UFP compatible 'UltraFLEXplus
        If (LCase(Pin_Type) = "setupfv") Then
            If glb_TesterType = "UltraFLEXplus" Then
                Dim arg_ary() As String
                Dim Vprog As Double
                arg_ary = Split(Pin_info_temp(2), ",")
                Vprog = CDbl(Spec_Evaluate_DC(arg_ary(0)))
                Force_Condition_V Pin_info_temp(0), Vprog
            Else
                'PAD_MTR_ANALOG_TEST_P:SetupFV:Vprog,Irange,CustomizeWaitTime
                Call SetupDCVI_ForceV(Pin_info_temp(0), Pin_info_temp(2))
            End If
            GoTo continue1
        End If
        
        If (LCase(Pin_Type) = "setupfi") Then
            If glb_TesterType = "UltraFLEXplus" Then
            Else
                'PAD_MTR_ANALOG_TEST_P:SetupFI:Vprog,Iprog
                ' Mode Alarm: voltage above Vprog
                ' Voltage Clamp Alarm: voltage above Vprog+960mV
                Call SetupDCVI_ForceI(Pin_info_temp(0), Pin_info_temp(2))
            End If
            GoTo continue1
        End If
        
        If (LCase(Pin_value) = "restoredcvi") Then
            If glb_TesterType = "UltraFLEXplus" Then
                With TheHdw.DCVS.Pins(Pin_info_temp(0))
                    .Gate = False
                    .Disconnect
                End With
            Else
                With TheHdw.DCVI.Pins(Pin_info_temp(0))
                    .Gate = False
                    .Disconnect
                End With
            End If
            GoTo continue1
        End If
        '' for UFP_Corr fix 200409
        If (LCase(Pin_value) = "disconnectdcvs") Then
            With TheHdw.DCVS.Pins(Pin_info_temp(0))
                .mode = tlDCVSModeHighImpedance   'Fix Mode alarm issue by cs 20201020
                .Voltage.value = 0
                .Gate = False
                .Meter.mode = tlDCVSMeterVoltage '20201006 CT add to fix MTRGR CZ error  "ERROR DCVS:0074 : DIB connect at force current mode is not allowed. Please set to high impedance or force voltage mode prior DIB connect. "
                .Disconnect
            End With
            GoTo continue1
        End If
        If (LCase(Pin_Type) = "vmode") Then
            With TheHdw.DCVS.Pins(PinName)
'                .mode = pin_value
                
                If (PrePatStore = True) Then
                    If (PrePat_Restore_String = "") Then
                        PrePat_Restore_String = PinName + ":vmode:" + CStr(.mode)
                    Else
                        PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":vmode:" + CStr(.mode)
                    End If
                ElseIf (PreMeasStore = True) Then
                    If (PreMeas_Restore_String = "") Then
                        PreMeas_Restore_String = PinName + ":vmode:" + CStr(.mode)
                    Else
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":vmode:" + CStr(.mode)
                    End If
                Else
                'Do nothing
                End If
                
                .mode = Pin_value
                
            End With
            GoTo continue1
        End If
''===============================================================================================
''170425 update for trace compensation
        If (LCase(Pin_Type) = "term") Then
           GoTo continue1
        End If

        If (UCase(Pin_Type) = "I") And pin_value <> "" Then
            ' Support ForceI condition of DCVI and DCVS in UF/UFP
            ' Use Format--> PinName:I:ForceI_Value,V_clamp_Range (DCVS/DCVI)
            ' Use Format--> PinName:I:ForceI_Value,V_clampH & V_ClampL (UP1600/UP2200)
            '==== Pin Norestore function for DCVS and DCVI ==== '20221026
            If NoRestore_Flag = 1 Then
                PrePatStore = False
                'Force_Condition_I PinName, CDbl(pin_value)
                Call Force_Condition_I(Pin_info_temp(0), Pin_info_temp(2))
                PrePatStore = True
            ElseIf NoRestore_Flag = 2 Then
                PreMeasStore = False
                'Force_Condition_I PinName, CDbl(pin_value)
                Call Force_Condition_I(Pin_info_temp(0), Pin_info_temp(2))
                PreMeasStore = True
            Else
                'Force_Condition_I PinName, CDbl(pin_value)
                Call Force_Condition_I(Pin_info_temp(0), Pin_info_temp(2))
            End If
            GoTo continue1
        End If
        '///////////////////////////////// Case V* ///////////////////////////////////////////////////////
        If (UCase(Pin_Type) = "V") And pin_value <> "" Then
            ' Support ForceV condition of DCVI and DCVS in UF/UFP
            ' Use Format--> PIN_NAME:V:VProg[@Ramping@Step@[Interval_Time]],[Irange],[CustomizeWaitTime]
            '==== Pin Norestore function for DCVS and DCVI ==== '20221026
            If NoRestore_Flag = 1 Then
                PrePatStore = False
                'Force_Condition_V PinName, CDbl(pin_value), pin_restore_analog
                Call Force_Condition_V(Pin_info_temp(0), Pin_info_temp(2), pin_restore_analog)
                PrePatStore = True
            ElseIf NoRestore_Flag = 2 Then
                PreMeasStore = False
                'Force_Condition_V PinName, CDbl(pin_value), pin_restore_analog
                Call Force_Condition_V(Pin_info_temp(0), Pin_info_temp(2), pin_restore_analog)
                PreMeasStore = True
            Else
				'[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
				If UCase(Pin_info_temp(2)) = "SHMOO_GLB" Then
					SD_Shmoo_GLB_Val = TheExec.Specs.Globals("SHMOO_GLB").CurrentValue
					'Force_Condition_V PinName, -999, pin_restore_analog
					Call Force_Condition_V(Pin_info_temp(0), -999, pin_restore_analog)
				Else
					'Force_Condition_V PinName, CDbl(pin_value), pin_restore_analog
					Call Force_Condition_V(Pin_info_temp(0), Pin_info_temp(2), pin_restore_analog)
				End If
            End If
            GoTo continue1
        End If
        If (UCase(Pin_Type) = "VIDS") And pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VIDS:" + Format(CStr(TheHdw.DCVS.Pins(PinName).Voltage.Main), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VIDS:" + Format(CStr(TheHdw.DCVS.Pins(PinName).Voltage.Main), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                Dim ids_p_ary() As String, ids_p_cnt As Long
                Dim ids_var As Variant
                TheExec.DataManager.DecomposePinList PinName, ids_p_ary, ids_p_cnt
                PreMeas_Restore_String = ids_var + ":VIDS:" + "RestoreByApplyLevelsTiming"
'                For Each ids_var In ids_p_ary
'                    If (PreMeas_Restore_String = "") Then
'                        PreMeas_Restore_String = ids_var + ":VIDS:" + Format(CStr(TheHdw.DCVS.Pins(ids_var).Voltage.Main), "0.000")
'                    Else
'                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + ids_var + ":VIDS:" + Format(CStr(TheHdw.DCVS.Pins(ids_var).Voltage.Main), "0.000")
'                    End If
'                Next ids_var
                TheHdw.DCVS.Pins(PinName).Voltage.Main = CDbl(Pin_value)
            ElseIf (PreMeasStore = False) Then
                    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
                    Exit For
            Else
            'Do nothing
            End If
            GoTo continue1
        End If
        If (UCase(Pin_Type) = "VALT") And Pin_value <> "" Then
            'Force_Condition_V PinName, CDbl(pin_value), pin_restore_analog
            SetPowerValue_Valt PinName, Pin_value
            GoTo continue1
        End If


        '///////////////////////////////// Case Range*for DCVS ///////////////////////////////////////////////////////
        If (UCase(Pin_Type) = "RANGE") And Pin_value <> "" Then
            If (PrePatStore = True) Then PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":Range:" + CStr(TheHdw.DCVS.Pins(PinName).CurrentRange.value)
            If mid(PrePat_Restore_String, 1, 1) = ";" Then PrePat_Restore_String = mid(PrePat_Restore_String, 2)
            If (PreMeasStore = True) Then PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":Range:" + CStr(TheHdw.DCVS.Pins(PinName).CurrentRange.value)
            If mid(PreMeas_Restore_String, 1, 1) = ";" Then PreMeas_Restore_String = mid(PreMeas_Restore_String, 2)
            TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
            TheHdw.DCVS.Pins(PinName).SetCurrentRanges CDbl(Pin_value), CDbl(Pin_value)
            TheHdw.DCVS.Pins(PinName).Gate = True
            TheExec.Datalog.WriteComment (TheExec.DataManager.instancename & " =====> Curr_meas Meter I range setting, " & PinName & " =" & CStr(TheHdw.DCVS.Pins(PinName).CurrentRange.value))
            GoTo continue1
        End If
        
        
        Dim i As Long
        'Modify for force condition "VRET" 20171213
        If (UCase(Pin_Type) = "VRET") And Pin_value <> "" Then
            Dim rp_ary() As String, rp_cnt As Long, pn As String, p_val As String
            TheExec.DataManager.DecomposePinList PinName, rp_ary, rp_cnt
            pn = Join(rp_ary, ",")
            p_val = Pin_value
            If UBound(rp_ary) >= 1 Then For i = 1 To UBound(rp_ary): p_val = p_val & "," & Pin_value: Next i
            If g_Retention_VDD = "" Then
                g_Retention_VDD = pn
                g_Retention_ForceV = p_val
            Else
                g_Retention_VDD = g_Retention_VDD & "," & pn
                g_Retention_ForceV = g_Retention_ForceV & "," & p_val
            End If
            GoTo continue1
        End If
        If (UCase(Pin_Type) = "VID") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VID:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVid)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VID:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVid)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VID:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVid)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VID:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVid)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVid) = CDbl(Pin_value)
            GoTo continue1
        End If
        '////////
        If (UCase(Pin_Type) = "VOD") And Pin_value <> "" Then
            If (LCase(TheExec.DataManager.PinType(PinName)) <> "differential") Then
                    If isDebugMode = True Then TheExec.AddOutput "[Alarm] Type: VOD ,Pin: " & PinName & " is not Differential Pin"
                    GoTo continue1
            End If
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VOD:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVod)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VOD:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVod)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VOD:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVod)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VOD:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVod)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVod) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        If (UCase(Pin_Type) = "VICM") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VICM:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVicm)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VICM:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVicm)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VICM:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVicm)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VICM:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVicm)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chVicm) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        ''''Add from S5E for AUSTXA
        If (UCase(Pin_Type) = "DIFFVT") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":DIFFVT:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chDiff_Vt)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":DIFFVT:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chDiff_Vt)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":DIFFVT:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chDiff_Vt)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":DIFFVT:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chDiff_Vt)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).DifferentialLevels.value(chDiff_Vt) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        If (UCase(Pin_Type) = "VIH") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VIH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVih)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VIH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVih)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VIH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVih)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VIH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVih)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.value(chVih) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        If (UCase(Pin_Type) = "VIL") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VIL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVil)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VIL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVil)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VIL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVil)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VIL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVil)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.value(chVil) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        If (UCase(Pin_Type) = "VOH") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVoh)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVoh)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVoh)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVoh)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.value(chVoh) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        If (UCase(Pin_Type) = "VOL") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVol)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVol)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVol)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVol)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.value(chVol) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        'IO_DS_TTR_200413
        If (UCase(Pin_Type) = "IOH") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":IOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chIoh)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":IOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chIoh)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":IOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chIoh)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":IOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chIoh)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.value(chIoh) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        If (UCase(Pin_Type) = "IOL") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":IOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chIol)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":IOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chIol)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":IOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chIol)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":IOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chIol)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.value(chIol) = CDbl(Pin_value)
            GoTo continue1
        End If
        'IO_DS_TTR_200413
        '/////////
        If (UCase(Pin_Type) = "VT") And Pin_value <> "" And (Not (UCase(TheExec.CurrentChanMap) Like "*FT*" And UCase(PinName) = "DDRIOPINS")) Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                        PrePat_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    Else
                        PrePat_Restore_String = PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    End If
                Else
                    If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                        PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    Else
                        PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    End If
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                        PreMeas_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    Else
                        PreMeas_Restore_String = PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    End If
                Else
                    If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    Else
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    End If
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeVt
            '[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
            If UCase(Pin_info_temp(2)) = "SHMOO_GLB" Then
                SD_Shmoo_GLB_Val = TheExec.Specs.Globals("SHMOO_GLB").CurrentValue
                For Each site In TheExec.sites
                    TheHdw.Digital.Pins(PinName).Levels.value(chVt) = CDbl(SD_Shmoo_GLB_Val)
                Next site
            Else
                TheHdw.Digital.Pins(PinName).Levels.value(chVt) = CDbl(pin_value)
            End If
            GoTo continue1
        ElseIf (UCase(TheExec.CurrentChanMap) Like "*FT*" And UCase(PinName) = "DDRIOPINS") Then
            PrePatStore = False
            GoTo continue1
        Else
        'Do nothing
        End If
        '/////////
        If (UCase(Pin_Type) = "HIZ") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                        PrePat_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    Else
                        PrePat_Restore_String = PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    End If
                Else
                    If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                        PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    Else
                        PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    End If
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                        PreMeas_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    Else
                        PreMeas_Restore_String = PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    End If
                Else
                    If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    Else
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVt)), "0.000")
                    End If
                End If
            Else
            'Do nothing
            End If
            
            '20210416, Add for UFP
            If glb_TesterType = "Jaguar" Then
                TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeHiZ
            Else
                TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeHiZ 'tlDriverModeLargeHiZ
            End If
 
            TheHdw.Digital.Pins(PinName).Levels.value(chVt) = CDbl(Pin_value)
            GoTo continue1
        End If

        If (UCase(Pin_Type) = "VCH") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VCH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVch)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VCH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVch)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VCH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVch)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VCH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVch)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.value(chVch) = CDbl(Pin_value)
            GoTo continue1
        End If
        '/////////
        If (UCase(Pin_Type) = "VCL") And Pin_value <> "" Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":VCL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVcl)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VCL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVcl)), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":VCL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVcl)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VCL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.value(chVcl)), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.Digital.Pins(PinName).Levels.value(chVcl) = CDbl(Pin_value)
            GoTo continue1
        End If
        '///////// 20180702 add for change D0, D1, D2, D3, R0, R1
         If (UCase(Pin_Type) Like "*D0*") And Pin_value <> "" Then
            Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0)), "0.000#########")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0)), "0.000#########")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0)), "0.000#########")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0)), "0.000#########")
                End If
            Else
            'Do nothing
            End If
                   TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0) = CDbl(Pin_value)
            GoTo continue1
        End If
        
         If (UCase(Pin_Type) Like "*D1*") And Pin_value <> "" Then
            Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1)), "0.000#########")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1)), "0.000#########")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1)), "0.000#########")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1)), "0.000#########")
                End If
            Else
            'Do nothing
            End If
                   TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1) = CDbl(Pin_value)
            GoTo continue1
        End If
        
        If (UCase(Pin_Type) Like "*D2*") And Pin_value <> "" Then
            Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2)), "0.000#########")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2)), "0.000#########")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2)), "0.000#########")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2)), "0.000#########")
                End If
            Else
            'Do nothing
            End If
                   TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2) = CDbl(Pin_value)
            GoTo continue1
        End If
        
        If (UCase(Pin_Type) Like "*D3*") And Pin_value <> "" Then
            Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3)), "0.000#########")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3)), "0.000#########")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3)), "0.000#########")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3)), "0.000#########")
                End If
            Else
            'Do nothing
            End If
                   TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3) = CDbl(Pin_value)
            GoTo continue1
        End If
        
        If (UCase(Pin_Type) Like "*R0*") And Pin_value <> "" Then
            Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0)), "0.000#########")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0)), "0.000#########")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0)), "0.000#########")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0)), "0.000#########")
                End If
            Else
            'Do nothing
            End If
                   TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0) = CDbl(Pin_value)
            GoTo continue1
        End If
        
        If (UCase(Pin_Type) Like "*R1*") And Pin_value <> "" Then
            Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1)), "0.000#########")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1)), "0.000#########")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1)), "0.000#########")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1)), "0.000#########")
                End If
            Else
            'Do nothing
            End If
                   TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1) = CDbl(Pin_value)
            GoTo continue1
        End If
        
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PPMU Connect Control\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If (LCase(Pin_info_temp(1)) = "connectppmu") Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":DisConnectPPMU"
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":DisConnectPPMU"
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":DisConnectPPMU"
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":DisConnectPPMU"
                End If
            Else
            'Do nothing
            End If
            
            TheHdw.PPMU.Pins(Pin_info_temp(0)).Connect
            GoTo continue1
        End If
        If (LCase(Pin_info_temp(1)) = "disconnectppmu") Then
            TheHdw.PPMU.Pins(Pin_info_temp(0)).Disconnect
            GoTo continue1
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Relay Control\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim OriginalRelayState As Long
        Dim RestoreString As String
        If (LCase(Pin_info_temp(1)) Like "*relay*") Then
            For Each site In TheExec.sites
                OriginalRelayState = TheHdw.Utility.Pins(PinName).States(tlUBStateProgrammed)
                Exit For
            Next site
            If (LCase(Pin_info_temp(1)) = "relay_on") Then
                TheHdw.Utility.Pins(PinName).State = tlUtilBitOn
            ElseIf (LCase(Pin_info_temp(1)) = "relay_off") Then
                TheHdw.Utility.Pins(PinName).State = tlUtilBitOff
            Else
            'Do nothing
            End If
            If UBound(Pin_info_temp) > 1 Then
                If LCase(Pin_info_temp(2)) Like "norestore" Then
                End If
            Else
                If OriginalRelayState = 0 Then
                    RestoreString = "relay_off"
                ElseIf OriginalRelayState = 1 Then
                    RestoreString = "relay_on"
                Else
                'Do nothing
                End If
                
                If (PrePatStore = True) Then
                    If (PrePat_Restore_String = "") Then
                        PrePat_Restore_String = PinName + ":" + RestoreString
                    Else
                        PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + RestoreString
                    End If
                ElseIf (PreMeasStore = True) Then
                    If (PreMeas_Restore_String = "") Then
                        PreMeas_Restore_String = PinName + ":" + RestoreString
                    Else
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + RestoreString
                    End If
                Else
                'Do nothing
                End If
            End If
            GoTo continue1
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Digital Connect Control\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If (LCase(Pin_info_temp(1)) = "connectdigital") Then
            TheHdw.Digital.Pins(Pin_info_temp(0)).Connect
            
            GoTo continue1
        End If
        If (LCase(Pin_info_temp(1)) = "disconnectdigital") Then
            TheHdw.Digital.Pins(Pin_info_temp(0)).Disconnect
            
            
            If (PrePatStore = True) Then  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ This part is for PLL I measurement 20170714 Kim
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":ConnectDigital"
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":ConnectDigital"
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":ConnectDigital"
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":ConnectDigital"
                End If
            Else
            'Do nothing
            End If
            
            GoTo continue1
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Digital compare\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If (LCase(Pin_info_temp(1)) = "disablecompare") Then
            TheHdw.Digital.Pins(Pin_info_temp(0)).DisableCompare = True
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":EnableCompare"
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":EnableCompare"
                End If
            End If
            GoTo continue1
        End If
        If (LCase(Pin_info_temp(1)) = "enablecompare") Then
            TheHdw.Digital.Pins(Pin_info_temp(0)).DisableCompare = False

            
            GoTo continue1
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If (InStr(LCase(Pin_info_temp(1)), "init") > 0) Then
            If (LCase(Pin_info_temp(1)) = "inithi") Then
                TheHdw.Digital.Pins(Pin_info_temp(0)).InitState = chInitHi
            End If
            If (LCase(Pin_info_temp(1)) = "initlo") Then
                TheHdw.Digital.Pins(Pin_info_temp(0)).InitState = chInitLo
            End If
            
            If (LCase(Pin_info_temp(1)) = "inithiz") Then
                TheHdw.Digital.Pins(Pin_info_temp(0)).InitState = chInitoff
            End If
            GoTo continue1
        End If
        
        
'        \\\\\\Setup Timeing for run pattern\\\\\\
        If (LCase(Pin_info_temp(1)) = "acspec") Then
            TheExec.Overlays.ApplyUniformSpecToHW Pin_info_temp(0), CDbl(Spec_Evaluate_AC(Pin_info_temp(2)))
            GoTo continue1
        End If
        If (LCase(Pin_info_temp(1)) = "tck") Then
            TheExec.Overlays.ApplyUniformSpecToHW Pin_info_temp(0), CDbl(Spec_Evaluate_AC(Pin_info_temp(2)))
            GoTo continue1
        End If
        If (LCase(Pin_info_temp(1)) = "shiftin") Then
            TheExec.Overlays.ApplyUniformSpecToHW Pin_info_temp(0), CDbl(Spec_Evaluate_AC(Pin_info_temp(2)))
            GoTo continue1
        End If
        
        
'\\\\\\Setup Timeing for nwire\\\\\\
'\\\\\\Disable FRC\\\\\\
        If (LCase(Pin_info_temp(1)) = "disable_frc") Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":enable_frc"
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":enable_frc"
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":enable_frc"
                Else
                    PreMeas_Restore_String = PrePat_Restore_String + ";" + PinName + ":enable_frc"
                End If
            Else
            'Do nothing
            End If
                Disable_FRC Pin_info_temp(0)
            GoTo continue1
        End If
        
'\\\\\\Enable FRC\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If (LCase(Pin_info_temp(1)) = "enable_frc") Then
            Enable_FRC Pin_info_temp(0)
            GoTo continue1
        End If

'\\\\\\Disable FRC\\\\\\ 20170817 disable FRC by switch relay
        If (LCase(Pin_info_temp(1)) = "disable_frc_relay") Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = PinName + ":enable_frc_relay"
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":enable_frc_relay"
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = PinName + ":enable_frc_relay"
                Else
                    PreMeas_Restore_String = PrePat_Restore_String + ";" + PinName + ":enable_frc_relay"
                End If
            Else
            'Do nothing
            End If

            TheExec.Datalog.WriteComment "=======Disable XO0======="
            TheHdw.Digital.Pins(Pin_info_temp(0)).Disconnect
            With TheHdw.PPMU.Pins(Pin_info_temp(0))
                .Disconnect
                .ForceV 0, 0.002
                .Connect
                .Gate = tlOn
            End With
            'TheHdw.Utility.Pins("K1").State = tlUtilBitOn
            GoTo continue1
        End If
        
'\\\\\\Enable FRC\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\20170817 enable FRC by switch relay
        If (LCase(Pin_info_temp(1)) = "enable_frc_relay") Then
            TheExec.Datalog.WriteComment "=======Enable XO0======="
            TheHdw.PPMU.Pins(Pin_info_temp(0)).Gate = tlOff
            TheHdw.PPMU.Pins(Pin_info_temp(0)).Disconnect
            TheHdw.Digital.Pins(Pin_info_temp(0)).Connect
            'TheHdw.Utility.Pins("K1").State = tlUtilBitOff
            GoTo continue1
        End If
        
        If isDebugMode = True Then TheExec.AddOutput "[Warning] Setup string" & cat_split_temp & " not support"
        
        
continue1:
    
    Next cat_split_temp
    
    ''20200903 Solve dsgrings alarm issue, do voltage ramping
    If ins_name Like "DSGRINGS_RNGV*" Or ins_name Like "PCMRINGS_*" Then
        Call HIPRampApplyLevel_AutoReadingContext(, , , TheExec.DataManager.instancename)
    End If

    If Not g_Vbump_function = True Then
       Shmoo_Set_Current_Point 'restore shmoo conidtion overrided in SetForceCondition
    End If
    
    If (gTerm_cond_All <> "") Then
        Call Trace_Compensation(flag_shmoo_set_current_point)
    End If
    TheHdw.Wait 0.001 '20160302  add settling time
    If (PrePatStore = True) Then
        TheExec.Datalog.WriteComment "Save PrePat Force Condtion:" & PrePat_Restore_String
    ElseIf (PreMeasStore = True) Then
        TheExec.Datalog.WriteComment "Save PreMeasForce Condtion:" & PreMeas_Restore_String
    Else
    'Do nothing
    End If
''================================================================================================================================

    PrePatStore = False
    PreMeasStore = False
    
    '/////////////////// 20180703 add for clean content of  Charz_Force_Power_condition ////////////////////////
            If TheExec.DevChar.Setups.IsRunning = True Then
                Dim SetupName As String
                Dim X_RangeFrom As Double
                Dim Y_RangeFrom As Double
                
                SetupName = TheExec.DevChar.Setups.ActiveSetupName
                If Not ((TheExec.DevChar.Results(SetupName).StartTime Like "1/1/0001*" Or TheExec.DevChar.Results(SetupName).StartTime Like "0001/1/1*")) Then
                    With TheExec.DevChar.Setups(SetupName)
                        If .Shmoo.axes.Count > 1 Then
                            X_RangeFrom = .Shmoo.axes(tlDevCharShmooAxis_X).Parameter.range.from
                            Y_RangeFrom = .Shmoo.axes(tlDevCharShmooAxis_Y).Parameter.range.from
                            For Each site In TheExec.sites ''20181101 current point need site value
                                XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.axes(tlDevCharShmooAxis_X).value
                                YVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.axes(tlDevCharShmooAxis_Y).value
                            Next site
                            If XVal = X_RangeFrom And YVal = Y_RangeFrom Then
                                gl_flag_end_shmoo = False
                            End If
                            If gl_flag_end_shmoo = True Then
                                Charz_Force_Power_condition = vbNullString
                            End If
                        Else
                            X_RangeFrom = .Shmoo.axes(tlDevCharShmooAxis_X).Parameter.range.from
                            For Each site In TheExec.sites ''20181101 current point need site value
                                XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.axes(tlDevCharShmooAxis_X).value
                            Next site
                            If XVal = X_RangeFrom Then
                                gl_flag_end_shmoo = False
                            End If
                            If gl_flag_end_shmoo = True Then
                                Charz_Force_Power_condition = vbNullString
                            End If
                        End If
                    End With
                End If
            End If
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    
    
    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "SetForceCondition") 'Add ErrHandler 2023/05/29
    If isDebugMode = True Then TheExec.AddOutput "Error ForeceCondition : " & Setup_string
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function HIPRampApplyLevel_AutoReadingContext(Optional ByVal ApplyPins As String = "CorePower", Optional RampingStep As Double = 1, Optional RampWaitTime As Double = 0, Optional instancename As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    ''SWLINZA20171120, for ramping voltage for each ATPG and Mbist instance

    Dim Apply_Pins_Ary() As String
    Dim Apply_Pins_count As Long
    Dim Extra_RampingTime As Double: Extra_RampingTime = RampWaitTime 'RampDown_Time = 0
    'Dim RampingStep As Double
    Dim Original_voltage() As Double
    Dim Apply_TargetVoltage() As Double
    Dim DiffVoltage() As Double
    Dim RampingVoltage() As Double
    Dim Voltage_from_HW As String
    Dim i, j As Integer
    Dim Current_DCCategory As String
    Dim Current_DCSelector As String
    Dim TestBlock As String
    Dim SepcSymbolic As String
    Dim ApplyPins_String As String
    Dim ApplyPins_Boolean() As Boolean
    'Dim AllPins_needApply As Boolean
    Dim Dummy_tempStr As String
    
''    If TheExec.EnableWord("Ramping_MbistATPG") = False Then Exit Function
    If Not UCase(TheExec.CurrentJob) Like "CP*" Then
            ApplyPins = Replace(UCase(ApplyPins), "COREPOWER", "COREPOWER_FT")
    End If
    
    TheExec.DataManager.DecomposePinList ApplyPins, Apply_Pins_Ary(), Apply_Pins_count
    ReDim Original_voltage(Apply_Pins_count - 1) As Double
    ReDim DiffVoltage(Apply_Pins_count - 1) As Double
    ReDim RampingVoltage(Apply_Pins_count - 1) As Double
    ReDim Apply_TargetVoltage(Apply_Pins_count - 1) As Double
    ReDim ApplyPins_Boolean(Apply_Pins_count - 1) As Boolean
    
    '----- to get target voltage from DC spec for each instance -----
    'Apply_TargetVoltage
    
    'Swlinza 20180126, to save test time in IGXL9.0, use this command instead of following two
    TheExec.DataManager.GetInstanceContext Current_DCCategory, Current_DCSelector, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr
    'Current_DCCategory = TheExec.TestInstances.Item(InstanceName).TimingAndLevels.DCCategory
    'Current_DCSelector = TheExec.TestInstances.Item(InstanceName).TimingAndLevels.DCSelector
    
'    If Current_DCCategory = Previous_DCCategory And Current_DCSelector = Previous_DCSelector Then
'        Exit Function
'    Else
        Previous_DCCategory = Current_DCCategory
        Previous_DCSelector = Current_DCSelector
'    End If
    
    TestBlock = mid(Current_DCCategory, 1, 3)
    
    Select Case UCase(TestBlock)
        Case UCase("Soc")
            SepcSymbolic = "_VAR_S"
        Case UCase("Cpu")
            SepcSymbolic = "_VAR_C"
        Case UCase("Gfx")
            SepcSymbolic = "_VAR_G"
        Case UCase("RTO")
            SepcSymbolic = "_VAR_R"
        Case Else
            SepcSymbolic = "_VAR"
    End Select
    
    Dim pin As Double
    '------ to calculate ramping voltage for each pins ------
    'AllPins_needApply = False
    If f_restorePrePat Then
        For i = 0 To Apply_Pins_count - 1
            Original_voltage(i) = FormatNumber(TheHdw.DCVS.Pins(Apply_Pins_Ary(i)).Voltage.Alt, 3)
            Apply_TargetVoltage(i) = TheExec.Specs.DC.item(Apply_Pins_Ary(i) & SepcSymbolic).Categories.item(Current_DCCategory).Selectors.item(Current_DCSelector).ContextValue
            DiffVoltage(i) = Original_voltage(i) - Apply_TargetVoltage(i)
            RampingVoltage(i) = FormatNumber((DiffVoltage(i) / RampingStep), 3)
            If Apply_TargetVoltage(i) = Original_voltage(i) Or Abs(DiffVoltage(i)) < 0.001 * RampingStep Then
                ApplyPins_Boolean(i) = False
            Else
                ApplyPins_Boolean(i) = True
                'AllPins_needApply = True
            End If
        Next i
        'If AllPins_needApply = False Then Exit Function
        '--------- Ramp down for retention voltage ------'
        For i = 0 To RampingStep - 1
            For j = 0 To Apply_Pins_count - 1
                If ApplyPins_Boolean(j) = True Then
                    If i = RampingStep - 1 Then
                        TheHdw.DCVS.Pins(Apply_Pins_Ary(j)).Voltage.Main = Apply_TargetVoltage(j)
                    Else
                        TheHdw.DCVS.Pins(Apply_Pins_Ary(j)).Voltage.Main = Original_voltage(j) - RampingVoltage(j) * (i + 1)
                    End If
                End If
            Next j
            TheHdw.DCVS.Pins(ApplyPins).Voltage.Output = tlDCVSVoltageMain
            For pin = 0 To UBound(Apply_Pins_Ary) - 1
                TheHdw.DCVS.Pins(Apply_Pins_Ary(pin)).Voltage.Alt = TheHdw.DCVS.Pins(Apply_Pins_Ary(pin)).Voltage.Main
            Next
            TheHdw.DCVS.Pins(ApplyPins).Voltage.Output = tlDCVSVoltageAlt
    '        thehdw.Wait Extra_RampingTime / RampingStep
        Next i
        TheHdw.DCVS.Pins(ApplyPins).Voltage.Output = tlDCVSVoltageMain
        f_restorePrePat = False
    Else
        For i = 0 To Apply_Pins_count - 1
            Original_voltage(i) = FormatNumber(TheHdw.DCVS.Pins(Apply_Pins_Ary(i)).Voltage.Main, 3)
            Apply_TargetVoltage(i) = TheHdw.DCVS.Pins(Apply_Pins_Ary(i)).Voltage.Alt.value
            DiffVoltage(i) = Original_voltage(i) - Apply_TargetVoltage(i)
            RampingVoltage(i) = FormatNumber((DiffVoltage(i) / RampingStep), 3)
            If Apply_TargetVoltage(i) = Original_voltage(i) Or Abs(DiffVoltage(i)) < 0.001 * RampingStep Then
                ApplyPins_Boolean(i) = False
            Else
                ApplyPins_Boolean(i) = True
                'AllPins_needApply = True
            End If
        Next i
        'If AllPins_needApply = False Then Exit Function
        '--------- Ramp down for retention voltage ------'
        For i = 0 To RampingStep - 1
            For j = 0 To Apply_Pins_count - 1
                If ApplyPins_Boolean(j) = True Then
                    If i = RampingStep - 1 Then
                        TheHdw.DCVS.Pins(Apply_Pins_Ary(j)).Voltage.Alt = Apply_TargetVoltage(j)
                    Else
                        TheHdw.DCVS.Pins(Apply_Pins_Ary(j)).Voltage.Alt = Original_voltage(j) - RampingVoltage(j) * (i + 1)
                    End If
                End If
            Next j
            TheHdw.DCVS.Pins(ApplyPins).Voltage.Output = tlDCVSVoltageAlt
            For pin = 0 To UBound(Apply_Pins_Ary) - 1
                TheHdw.DCVS.Pins(Apply_Pins_Ary(pin)).Voltage.Main = TheHdw.DCVS.Pins(Apply_Pins_Ary(pin)).Voltage.Alt
            Next
            
            TheHdw.DCVS.Pins(ApplyPins).Voltage.Output = tlDCVSVoltageMain
    '        thehdw.Wait Extra_RampingTime / RampingStep
        Next i
        TheHdw.DCVS.Pins(ApplyPins).Voltage.Output = tlDCVSVoltageAlt
        f_restorePrePat = True

    End If
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "HIPRampApplyLevel_AutoReadingContext") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Add_Term_Restore(Setup_string As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim condition_ary() As String
    Dim cond_element_ary() As String
    Dim cond As Variant
    Dim Pin_Cnt As Long, Pin_Ary() As String
    Dim p As String
    Dim term_cond_ary() As String
    Dim cond_add As String
    Dim term_setting_ary() As String, term_type As String
    Dim DevChar_Setup As String
    Dim p_temp() As String
    Dim i As Integer
    ' Setup_string = p1:term:S,0,50, => gTerm_Restore_cond = p1:vih:1;p1:vil:0
    If InStr(Setup_string, "term") <= 0 Then
        gTerm_cond_All = vbNullString
        Exit Function
    Else:
'        gTerm_cond_Flag = True
    End If
    If TheExec.DevChar.Setups.IsRunning = True Then
        DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
        If gTerm_Restore_cond <> "" And LCase(Setup_string) Like "*restore*" Then
            Setup_string = gTerm_Restore_cond & ";" & Setup_string
        End If
'        gTerm_cond_shm_Flag = True
        'Exit Function
        
    End If
    gTerm_cond_All = vbNullString             'p1:term:S,0,50,
    gTerm_Restore_cond = vbNullString         'p1:vih:1;p1:vil:0
    gTerm_cond_shm_Flag = False
    condition_ary = Split(Setup_string, ";")

    For Each cond In condition_ary
        If cond <> "" Then
            cond_element_ary = Split(cond, ":")
            p = cond_element_ary(0)
            If UBound(cond_element_ary) > 0 Then
                If cond_element_ary(1) = "term" Then
                    term_setting_ary = Split(cond_element_ary(2), ",")
                    term_type = term_setting_ary(0)
                    If gTerm_cond_All = "" Then
                        gTerm_cond_All = cond
                    Else:
                        gTerm_cond_All = gTerm_cond_All & ";" & cond
                    End If
                    If term_type = "s" Then
                        cond_add = p & ":vih:" & Format(TheHdw.Digital.Pins(p).Levels.value(chVih), "0.0000") & ";" & _
                                   p & ":vil:" & Format(TheHdw.Digital.Pins(p).Levels.value(chVil), "0.0000")
                    ElseIf term_type = "d" Then
                        p_temp = Split(p, ",")
                        For i = 0 To UBound(p_temp)
                            If (TheExec.DataManager.PinType(p_temp(i)) = "Differential") Then
                                cond_add = p_temp(i) & ":vid:" & Format(TheHdw.Digital.Pins(p_temp(i)).DifferentialLevels.value(chVid), "0.0000") & ";" & _
                                           p_temp(i) & ":vicm:" & Format(TheHdw.Digital.Pins(p_temp(i)).DifferentialLevels.value(chVicm), "0.0000")
                            Else
                                cond_add = p_temp(i) & ":vih:" & Format(TheHdw.Digital.Pins(p_temp(i)).Levels.value(chVih), "0.0000") & ";" & _
                                p_temp(i) & ":vil:" & Format(TheHdw.Digital.Pins(p_temp(i)).Levels.value(chVil), "0.0000")
                            End If
                        Next i
                    Else
                    'Do nothing
                    End If
                    If gTerm_Restore_cond = "" Then
                        gTerm_Restore_cond = cond_add
                    Else:
                        gTerm_Restore_cond = gTerm_Restore_cond & ";" & cond_add
                    End If
               End If
            End If
        End If
    Next cond
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Add_Term_Restore") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function





Public Function Trace_Compensation(flag_shmoo_set_current_point As Boolean) As Double
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim pin_info() As String
    Dim TypeName As String
    Dim Vi_dut, Rs, Vt_dut As Double    'Rt_dut_d, Rt_dut_u
    Dim PinName As String
    Dim PinData As String
    Dim Delta_Rs As New PinListData
    Dim site As Variant
    Dim Job As String
    Dim vi_h, vi_l, vi_d, vi_cm As New PinListData
    Dim pin_grp() As String, pin_grp1() As String
    Dim Pin_Cnt As Long, pin_cnt1 As Long, x As Long
    Dim pin_temp As Variant, p1 As Variant
    Dim Rak_val() As Double
    'Dim Rt_dut As New SiteDouble
    Dim Rt_dut As New PinListData
    Dim Term_cond As Variant, term_cond_ary() As String
    Dim PinData_ary() As String
    Dim DevChar_Setup As String
    Dim Shmoo_Tracking_Item As Variant
    Dim shmoo_axis As Variant
    Dim Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_value As Double, Port_name As String
    Dim Shmoo_Step_Name As String, Shmoo_TimeSets As String
    Dim Is_singleEnd As Boolean
    Dim Vicm_old, Vid_old, Vicm_new, Vid_new, Vih_new, Vil_new As Double
    Dim i As Integer
    Dim pin_temp_P_PPMUForceValue, pin_temp_N_PPMUForceValue, Vi_pgm_P, Vi_pgm_N As Double
    Dim pin_temp_N As String
    Dim Rt_dut_u As Double  'New SiteDouble
    Dim Rt_dut_d As Double  'New SiteDouble
    
    
    term_cond_ary = Split(gTerm_cond_All, ";")
    For Each Term_cond In term_cond_ary
        pin_info = Split(Term_cond, ":")
        PinName = pin_info(0)
        PinData = pin_info(2)
        Vt_dut = -999
        Rt_dut_u = 999999999
        Rt_dut_d = 999999999
        
        PinData_ary = Split(PinData, ",")
        TypeName = PinData_ary(0)
        
        If (PinData_ary(1) <> "") Then
            Vt_dut = CDbl(PinData_ary(1))
        End If
        If (PinData_ary(2) <> "") Then
            If IsNumeric(PinData_ary(2)) Then
                Rt_dut_u = CDbl(PinData_ary(2))
                Rt_dut = Rt_dut_u
            Else
                Rt_dut = GetStoreDataAllType(PinData_ary(2))
            End If
        End If
        
        Rs = 50
        If (TheExec.DataManager.PinType(Split(PinName, ",")(0)) = "Differential") Then
        '180430 avoid PinName is PinGroupPin
            Dim GroupDiff_pin_grp() As String
            Dim GruopDiff_Pin_Cnt As Long, iI As Long
            Dim GroupDiff_pin_temp As Variant
            Call TheExec.DataManager.DecomposePinList(PinName, GroupDiff_pin_grp, GruopDiff_Pin_Cnt)
            pin_grp = Split(PinName, ",")
            Pin_Cnt = UBound(pin_grp)
            Is_singleEnd = False
        Else:
            Call TheExec.DataManager.DecomposePinList(PinName, pin_grp, Pin_Cnt)
            Is_singleEnd = True
        End If
        
        
        For i = 0 To UBound(pin_grp)
        pin_temp = pin_grp(i)
            If (TypeName = "s") Then
                pin_temp = LCase(pin_temp)
                Delta_Rs.AddPin (pin_temp)

                For Each site In TheExec.sites
                    Rak_val = TheHdw.PPMU.ReadRakValuesByPinnames(pin_temp, site)
                    Delta_Rs.pin(pin_temp).value = CurrentJob_Card_RAK.Pins(pin_temp).value + Rak_val(0)
                Next site
                
                If (TheExec.DataManager.PinType(pin_temp) <> "I/O") Then
                    If isDebugMode = True Then TheExec.AddOutput "Pin : " + pin_temp + " Is Not Single-End Pin!"
                    Exit Function
                End If
                If (Vt_dut = -999) Then Vt_dut = 0
                For Each site In TheExec.sites
                    TheHdw.Digital.Pins(pin_temp).Levels.value(chVih) = TheHdw.Digital.Pins(pin_temp).Levels.value(chVih) * (1 + (Rs + Delta_Rs.pin(pin_temp).value(site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) - Vt_dut * (Rs + Delta_Rs.pin(pin_temp).value(site)) / Rt_dut_u
                    TheHdw.Digital.Pins(pin_temp).Levels.value(chVil) = TheHdw.Digital.Pins(pin_temp).Levels.value(chVil) * (1 + (Rs + Delta_Rs.pin(pin_temp).value(site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) - Vt_dut * (Rs + Delta_Rs.pin(pin_temp).value(site)) / Rt_dut_u
                    TheExec.Datalog.WriteComment pin_temp & ": VIH(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.value(chVih))
                    TheExec.Datalog.WriteComment pin_temp & ": VIL(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.value(chVil))
                Next site
            ElseIf (TypeName = "d") Then
                ' 180504 Assume the RAK value of Differential pin P/N is the same, just take P or N value
                '        if does not assumethe same, the formulas below needs to be modified.
                If Is_singleEnd = False Then
                    For iI = 0 To GruopDiff_Pin_Cnt - 2
                        Delta_Rs.AddPin LCase(GroupDiff_pin_grp(iI))
                        For Each site In TheExec.sites
                                Rak_val = TheHdw.PPMU.ReadRakValuesByPinnames(GroupDiff_pin_grp(iI), site)
                                Delta_Rs.pin(GroupDiff_pin_grp(iI)).value = CurrentJob_Card_RAK.Pins(GroupDiff_pin_grp(iI)).value - Rak_val(0)
                            If (Vt_dut <> -999) Then ''Vt with value
                                    TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVid) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVid) * (1 + (Rs + Delta_Rs.pin(GroupDiff_pin_grp(iI)).value(site)) * (1 / Rt_dut)) - Vt_dut * (Rs + Delta_Rs.pin(GroupDiff_pin_grp(iI)).value(site)) / Rt_dut
                                    TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVicm) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVicm) * (1 + (Rs + Delta_Rs.pin(GroupDiff_pin_grp(iI)).value(site)) * (1 / Rt_dut)) - Vt_dut * (Rs + Delta_Rs.pin(GroupDiff_pin_grp(iI)).value(site)) / Rt_dut
                                    TheExec.Datalog.WriteComment pin_temp & ": VID(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVid))
                                    TheExec.Datalog.WriteComment pin_temp & ": VICM(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVicm))
                            Else:
                                    TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVid) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVid) * (Rs + Delta_Rs.pin(GroupDiff_pin_grp(iI)).value + Rt_dut) / Rt_dut
                                    TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVicm) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVicm)
                                    TheExec.Datalog.WriteComment pin_temp & ": VID(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVid))
                                    TheExec.Datalog.WriteComment pin_temp & ": VICM(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.value(chVicm))
                            End If
                        Next site
                    Next iI
                Else
                    pin_temp = LCase(pin_temp)
                    Delta_Rs.AddPin (pin_temp)
                    'pin_temp_N = pin_grp(i + 1)
                    pin_temp_N = Replace(pin_temp, "p", "m")
                    For Each site In TheExec.sites
                        Rak_val = TheHdw.PPMU.ReadRakValuesByPinnames(pin_temp, site)
                        Delta_Rs.pin(pin_temp).value = CurrentJob_Card_RAK.Pins(pin_temp).value
                        
                        'Delta_Rs.Pin(pin_temp).Value = CurrentJob_Card_RAK.Pins(pin_temp).Value - RAK_Val(0)
                    Next site
                    
                    If (TheExec.DataManager.PinType(pin_temp) <> "I/O") Then
                        If isDebugMode = True Then TheExec.AddOutput "Pin : " + pin_temp + " Is Not Single-End Pin!"
                        Exit Function
                    End If
                    For Each site In TheExec.sites
                        If TheHdw.PPMU.Pins(pin_temp).Gate = tlOn Then
                            pin_temp_P_PPMUForceValue = TheHdw.PPMU.Pins(pin_temp).Voltage.value
                            pin_temp_N_PPMUForceValue = TheHdw.PPMU.Pins(pin_temp_N).Voltage.value
                            If Not (Vt_dut = -999) Then
                                Vi_pgm_P = Format(pin_temp_P_PPMUForceValue * (1 + (0 + CurrentJob_Card_RAK.Pins(pin_temp).value)) * (1 / Rt_dut), "0.000") ''0 is RS
                                Vi_pgm_N = Format(pin_temp_N_PPMUForceValue * (1 + (0 + CurrentJob_Card_RAK.Pins(pin_temp_N).value)) * (1 / Rt_dut), "0.000") ''0 is RS
                                
                                TheExec.Datalog.WriteComment "P_Pin : " & UCase(pin_temp) & " site:" & site & " Force_V_Value = " & Vi_pgm_P
                                TheExec.Datalog.WriteComment "N_Pin : " & UCase(pin_temp_N) & " site:" & site & " Force_V_Value = " & Vi_pgm_N
                            Else
                                Vi_pgm_P = Format(pin_temp_P_PPMUForceValue * (1 + (0 + CurrentJob_Card_RAK.Pins(pin_temp).value) / (2 * Rt_dut.Pins(pin_temp).value)) - pin_temp_N_PPMUForceValue * ((0 + CurrentJob_Card_RAK.Pins(pin_temp).value) / (2 * Rt_dut.Pins(pin_temp).value)), "0.000")  ''0 is RS
                                Vi_pgm_N = Format(pin_temp_N_PPMUForceValue * (1 + (0 + CurrentJob_Card_RAK.Pins(pin_temp_N).value) / (2 * Rt_dut.Pins(pin_temp).value)) - pin_temp_P_PPMUForceValue * ((0 + CurrentJob_Card_RAK.Pins(pin_temp_N).value) / (2 * Rt_dut.Pins(pin_temp).value)), "0.000") ''0 is RS
    
                                TheExec.Datalog.WriteComment "P_Pin : " & UCase(pin_temp) & " site:" & site & " Force_V_Value = " & Vi_pgm_P
                                TheExec.Datalog.WriteComment "N_Pin : " & UCase(pin_temp_N) & " site:" & site & " Force_V_Value = " & Vi_pgm_N
                            End If
                            Call TheHdw.PPMU.Pins(pin_temp).ForceV(Vi_pgm_P, 0.02)
                            Call TheHdw.PPMU.Pins(pin_temp_N).ForceV(Vi_pgm_N, 0.02)
                            'i = i + 1
                        Else
                            Vicm_old = (TheHdw.Digital.Pins(pin_temp).Levels.value(chVih) + TheHdw.Digital.Pins(pin_temp).Levels.value(chVil)) / 2
                            Vid_old = TheHdw.Digital.Pins(pin_temp).Levels.value(chVih) - TheHdw.Digital.Pins(pin_temp).Levels.value(chVil)
                            If Not (Vt_dut = -999) Then
                                    Vicm_new = Vicm_old * (1 + (Rs + Delta_Rs.pin(pin_temp).value(site)) * (1 / Rt_dut.Pins(pin_temp).value)) - Vt_dut * (Rs + Delta_Rs.pin(pin_temp).value(site)) / Rt_dut.Pins(pin_temp).value
                                    Vid_new = Vid_old * (1 + (Rs + Delta_Rs.pin(pin_temp).value(site)) * (1 / Rt_dut.Pins(pin_temp).value)) - Vt_dut * (Rs + Delta_Rs.pin(pin_temp).value(site)) / Rt_dut.Pins(pin_temp).value
                                    Vih_new = Vicm_new + Vid_new / 2
                                    Vil_new = Vicm_new - Vid_new / 2
    
    
                                    TheHdw.Digital.Pins(pin_temp).Levels.value(chVih) = Vih_new
                                    TheHdw.Digital.Pins(pin_temp).Levels.value(chVil) = Vil_new
                                    TheExec.Datalog.WriteComment pin_temp & ": VIH(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.value(chVih))
                                    TheExec.Datalog.WriteComment pin_temp & ": VIL(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.value(chVil))
                            Else
                                    Vicm_new = Vicm_old * (Rs + Delta_Rs.pin(pin_temp).value + Rt_dut.Pins(pin_temp).value) / Rt_dut.Pins(pin_temp).value
                                    Vid_new = Vid_old
                                    Vih_new = Vicm_new + Vid_new / 2
                                    Vil_new = Vicm_new - Vid_new / 2
                                    TheHdw.Digital.Pins(pin_temp).Levels.value(chVih) = Vih_new
                                    TheHdw.Digital.Pins(pin_temp).Levels.value(chVil) = Vil_new
                                    TheExec.Datalog.WriteComment pin_temp & ": VIH(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.value(chVih))
                                    TheExec.Datalog.WriteComment pin_temp & ": VIL(" & site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.value(chVil))
                            End If
                        End If
                    Next site
                End If
            Else
            'Do nothing
            End If
        Next i
    Next Term_cond

    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Trace_Compensation") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function SetPowerValue(ByVal pin As String, ByVal Pin_value As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim PinList() As String
    Dim PinNum As Integer
    Dim get_type As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var As Variant
    Dim PinName As String
    Dim site As Variant 'Carter, 20240304
    
    Call TheExec.DataManager.DecomposePinList(pin, PinList, typesCount)
    
    
    For Each var In PinList

        Call TheExec.DataManager.GetChannelTypes(var, typesCount, stringTypes)
        If (stringTypes(0) Like "DCVS*") Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.000")
                End If
                If TheHdw.DCVS.Pins(var).Gate = False Then PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "disconnectdcvs"
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.000")
                End If
                '20210416, add for ufp
                If TheHdw.DCVS.Pins(var).Gate = False Then PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "disconnectdcvs"
            Else
            'Do nothing
            End If
            If g_Vbump_function = True Then 'add for SelSram function
                 If Not g_CharInputString_Voltage_Dict.Exists(UCase(var)) = True Then
                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(Pin_value)
                 Else
                    g_CharInputString_Voltage_Dict.Remove (UCase(var))
                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(Pin_value)
                 End If
            Else
                '''Add for UFP
                If TheHdw.DCVS.Pins(var).Gate = False Then
                    With TheHdw.DCVS.Pins(var)
                        .Disconnect tlDCVSConnectDefault
                        .Meter.mode = tlDCVSMeterCurrent
                        .mode = tlDCVSModeVoltage
                        .SetCurrentRanges pc_Def_UVS256HP_Init_MeasCurrRange, pc_Def_UVS256HP_Init_MeasCurrRange
                        .Voltage.value = 0#
                        .Connect tlDCVSConnectDefault
                        .Gate = True
                        '[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
                        If CDbl(pin_value) = -999 Then
                            For Each site In TheExec.sites
                                .Voltage.Main.value = SD_Shmoo_GLB_Val
                            Next site
                        Else
                            .Voltage.Main.value = CDbl(pin_value)
                        End If
                    End With
                Else
					''20200903 Solve Bora dsgrings alarm issue
					''1. while store interpose PrePat, voltage output selector would be Vmain, save value and pin in Valt
					''2. while restore original voltage, voltage output selector would be Valt, save value and pin in Vmain
					
                    '[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
                    If CDbl(pin_value) = -999 Then
                        For Each site In TheExec.sites
                            If TheExec.DataManager.InstanceName Like "DSGRINGS_RNGV*" And Not (f_restorePrePat) Then
                                TheHdw.DCVS.Pins(var).Voltage.Alt.value = CDbl(SD_Shmoo_GLB_Val)
                            Else
                                TheHdw.DCVS.Pins(var).Voltage.Main.value = CDbl(SD_Shmoo_GLB_Val)
                            End If
                        Next site
                    Else
                        If TheExec.DataManager.InstanceName Like "DSGRINGS_RNGV*" And Not (f_restorePrePat) Then
                            TheHdw.DCVS.Pins(var).Voltage.Alt.value = CDbl(pin_value)
                        Else
                            TheHdw.DCVS.Pins(var).Voltage.Main.value = CDbl(pin_value)
                        End If
                    End If										
                End If
            End If
            
        End If
        If (stringTypes(0) Like "DCVI*") Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.000")
                End If
            Else
                'Do nothing
            End If

            If g_Vbump_function = True Then 'add for SelSram function
                If Not g_CharInputString_Voltage_Dict.Exists(UCase(var)) = True Then
                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(Pin_value)
                Else
                    g_CharInputString_Voltage_Dict.Remove (UCase(var))
                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(Pin_value)
                End If
            Else
                With TheHdw.DCVI.Pins(var)
					If .Gate = True Then
						If .mode <> tlDCVIModeVoltage Then
                            .mode = tlDCVIModeVoltage
                        End If
						'[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
						 If CDbl(pin_value) = -999 Then
							For Each site In TheExec.sites
								.Voltage = CDbl(SD_Shmoo_GLB_Val)
							Next site
						 Else
							.Voltage = CDbl(pin_value) ' MI_TestCond_UVI80(i).FV_Val
						 End If
                    Else
						If .mode <> tlDCVIModeVoltage Then
							.Gate = False
							.mode = tlDCVIModeVoltage
						End If
						'[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
						 If CDbl(pin_value) = -999 Then
							For Each site In TheExec.sites
								.Voltage = CDbl(SD_Shmoo_GLB_Val)
							Next site
						 Else
							.Voltage = CDbl(pin_value) ' MI_TestCond_UVI80(i).FV_Val
						 End If
						.VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
						.SetCurrentAndRange 0.2, 0.2 'MI_TestCond_UVI80(i).CurrentRange, MI_TestCond_UVI80(i).CurrentRange
						.Connect tlDCVIConnectDefault
						.Gate = True
					End If
				End With
            End If
        End If
cont:
    Next var
    
    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "SetPowerValue") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
'20170419 add for case I
Public Function SetPowerValue_I(ByVal pin As String, ByVal Pin_value As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim PinList() As String
    Dim PinNum As Integer
    Dim get_type As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var As Variant
    Dim PinName As String
    

    
    Call TheExec.DataManager.DecomposePinList(pin, PinList, typesCount)
    
    
    For Each var In PinList '
        Call TheExec.DataManager.GetChannelTypes(var, typesCount, stringTypes)
        
        If (stringTypes(0) Like "DCVI*") Then
            TheHdw.DCVI.Pins(var).mode = tlDCVIModeCurrent
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = var + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(var).Current), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(var).Current), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = var + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(var).Current), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(var).Current), "0.000")
                End If
            Else
            'Do nothing
            End If
            TheHdw.DCVI.Pins(var).Current = CDbl(Pin_value)
        ElseIf (stringTypes(0) Like "DCVS*") Then
            If glb_TesterType = "Jaguar" Then
                If isDebugMode = True Then TheExec.AddOutput "[Warning]  No force I mode for DCVS"
            Else
                If (PrePatStore = True) Then
                    If (Format(CStr(TheHdw.DCVS.Pins(var).CurrentLimit.Source.FoldLimit.level.value), "0.000") = 0) Or TheHdw.DCVS.Pins(var).Gate = False Then
                        PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "disconnectdcvs"
                    Else
                        If (PrePat_Restore_String = "") Then
                            PrePat_Restore_String = var + ":" + "I:" + Format(CStr(TheHdw.DCVS.Pins(var).CurrentLimit.Source.FoldLimit.level.value), "0.000")
                        Else
                            PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "I:" + Format(TheHdw.DCVS.Pins(var).CurrentLimit.Source.FoldLimit.level.value, "0.000")
                        End If
                    End If
                    
                ElseIf (PreMeasStore = True) Then
                    '' for UFP_Corr fix 200409
                    If (Format(CStr(TheHdw.DCVS.Pins(var).CurrentLimit.Source.FoldLimit.level.value), "0.000") = 0) Or TheHdw.DCVS.Pins(var).Gate = False Then
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "disconnectdcvs"
                    Else
                        If (PreMeas_Restore_String = "") Then
                            PreMeas_Restore_String = var + ":" + "I:" + Format(CStr(TheHdw.DCVS.Pins(var).CurrentLimit.Source.FoldLimit.level.value), "0.000")
                        Else
                            PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "I:" + Format(CStr(TheHdw.DCVS.Pins(var).CurrentLimit.Source.FoldLimit.level.value), "0.000")
                        End If
                    End If
                Else
                'Do nothing
                End If
                With TheHdw.DCVS.Pins(var)
                    If (CDbl(Pin_value) = 0) And .Gate <> True Then
                        .Disconnect
                        .mode = tlDCVSModeHighImpedance
                        .Meter.mode = tlDCVSMeterVoltage
                        .Gate = True
                        .Connect
                    Else
                    
                        If CDbl(Pin_value) > 0 Then
                            .Voltage.value = 1 + 0.5
                        Else
                            '---------------------------------------UFP_Corr fix 20200413
                            .VoltageRange.value = 5.5
                            '---------------------------------------UFP_Corr fix 20200413
                            .Voltage.value = -1 - 0.5
                        End If
                        
                        .CurrentRange.value = Abs(CDbl(Pin_value))
                        .CurrentLimit.Source.FoldLimit.level.value = Abs(CDbl(Pin_value))
                        .CurrentLimit.Sink.FoldLimit.level.value = Abs(CDbl(Pin_value))
                        .Meter.mode = tlDCVSMeterVoltage
                        .Gate = True
                        .Connect
                        '---------------------------------------UFP_Corr fix 20200413
                        .mode = tlDCVSModeCurrent
                        '---------------------------------------UFP_Corr fix 20200413
                    
                    End If
                End With
            End If
        Else
            If isDebugMode = True Then TheExec.AddOutput "[Warning] " & var & "not support force I"
        End If
cont:
    Next var
    
    Exit Function

errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "SetPowerValue_I") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function PrintCharSetup(ByVal Setup_string As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim Cat_temp() As String
Dim cat_split_temp As Variant

Dim Pin_info_temp() As String

Dim PinName, Pin_Type, Pin_value As String
Dim OutputString As String
Dim temp As String
Dim site As Variant




    OutputString = vbNullString
  
    Cat_temp = Split(Setup_string, ";")
    For Each cat_split_temp In Cat_temp
    
        Pin_info_temp = Split(cat_split_temp, ":")
        If UBound(Pin_info_temp) < 2 Then GoTo continue2
        If (Pin_info_temp(2) = "") Then GoTo continue2

        PinName = Pin_info_temp(0)
        Pin_Type = Pin_info_temp(1)
        Pin_value = Pin_info_temp(2)
        
        PinName = Replace(PinName, "+", ",")
        
        
        
        If (Pin_Type = "V") Then
            temp = PrintCharValue(PinName)
            If (OutputString = "") Then
                OutputString = temp
            Else
                OutputString = OutputString & "," & temp
            End If
            GoTo continue2
        End If
        
continue2:
    Next cat_split_temp
    
    PrintCharSetup = OutputString
    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "PrintCharSetup") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function PrintCharValue(ByVal pin As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim PinList() As String
    Dim PinNum As Integer
    Dim get_type As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var As Variant
    Dim OutputString As String
    
    OutputString = vbNullString
    Call TheExec.DataManager.DecomposePinList(pin, PinList, typesCount)
    
    
    For Each var In PinList
        Call TheExec.DataManager.GetChannelTypes(var, typesCount, stringTypes)
        If (stringTypes(0) Like "DCVS*") Then
            If (OutputString = "") Then
'20170120 change print format
                OutputString = var & ":V:" & Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.0000")
            Else
                OutputString = OutputString & "," & var & ":V:" & Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.0000")
            End If
        End If

        If (stringTypes(0) Like "DCVI*") Then
            If (OutputString = "") Then
'20170120 change print format
                OutputString = var & ":V:" & Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.0000")
            Else
                OutputString = OutputString & "," & var & ":V:" & Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.0000")
            End If

        End If

    Next var
    

    PrintCharValue = OutputString
    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "PrintCharValue") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function Spec_Evaluate_DC_for_flow_loop(ByVal pininfo As String, ByVal condition_info As String, ByVal temp_pin_info As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim temp_pin_name, temp_flowint_name, calc_info As String
    Dim temp_pininfo_arr() As String
    Dim i As Integer
    Dim fake_calc_info, outstring_ori, outstring_evaluated As String
    Dim flow_int_value As Double
    Dim site As Variant
    
    For Each site In TheExec.sites
     
	    outstring_ori = pininfo & ":" & condition_info & ":" & temp_pin_info
	    
	    If InStr(temp_pin_info, "+") > 0 Or InStr(temp_pin_info, "-") > 0 Or InStr(temp_pin_info, "*") > 0 Or InStr(temp_pin_info, "/") > 0 Then
	        calc_info = temp_pin_info
	        
	        temp_pin_info = Replace(temp_pin_info, "(", vbNullString)
	        temp_pin_info = Replace(temp_pin_info, ")", vbNullString)
	        temp_pin_info = Replace(temp_pin_info, "+", "~")
	        temp_pin_info = Replace(temp_pin_info, "-", "~")
	        temp_pin_info = Replace(temp_pin_info, "*", "~")
	        temp_pin_info = Replace(temp_pin_info, "/", "~")
	        
	        fake_calc_info = Replace(calc_info, "(", " ( ")
	        fake_calc_info = Replace(fake_calc_info, ")", " ) ")
	        fake_calc_info = Replace(fake_calc_info, "+", " + ")
	        fake_calc_info = Replace(fake_calc_info, "-", " - ")
	        fake_calc_info = Replace(fake_calc_info, "*", " * ")
	        fake_calc_info = Replace(fake_calc_info, "/", " / ")
	        fake_calc_info = " " & fake_calc_info & " "
	        
	        temp_pininfo_arr = Split(temp_pin_info, "~")
	        
	        For i = 0 To UBound(temp_pininfo_arr)
	            If InStr(temp_pininfo_arr(i), "_") <> 0 Then
	            
	                If InStr(LCase(temp_pininfo_arr(i)), "flow_int") <> 0 Then
	                flow_int_value = TheExec.Flow.var(temp_pininfo_arr(i)).value
	                calc_info = Replace(fake_calc_info, " " & temp_pininfo_arr(i) & " ", flow_int_value, , 1)
	                fake_calc_info = calc_info
	                Else
	                temp_pin_name = temp_pininfo_arr(i)
	                temp_pininfo_arr(i) = CStr(TheExec.Specs.DC.item(mid(temp_pininfo_arr(i), 2)).CurrentValue(site))
	                calc_info = Replace(fake_calc_info, " " & temp_pin_name & " ", temp_pininfo_arr(i), , 1)
	                fake_calc_info = calc_info
	                End If
	            
	            
	            End If
	        Next i
	        Spec_Evaluate_DC_for_flow_loop = CStr(Evaluate(calc_info))
	        outstring_evaluated = pininfo & ":" & condition_info & ":" & calc_info & "=" & CStr(Spec_Evaluate_DC_for_flow_loop)
	        TheExec.Datalog.WriteComment "Calculate_Result:" & Trim(outstring_evaluated)
	        
	    Else
	        If (InStr(temp_pin_info, "_") = 1) Then
	            If TheExec.Specs.DC.Contains(mid(temp_pin_info, 2)) Then
	                Spec_Evaluate_DC_for_flow_loop = CStr(TheExec.Specs.DC.item(mid(temp_pin_info, 2)).CurrentValue(site))
	                outstring_evaluated = pininfo & ":" & condition_info & "=" & CStr(Spec_Evaluate_DC_for_flow_loop)
	                TheExec.Datalog.WriteComment "Calculate_Result:" & Trim(outstring_evaluated)
	            Else
	                Spec_Evaluate_DC_for_flow_loop = CStr(temp_pin_info)
	            End If
	        Else
	            If InStr(LCase(temp_pin_info), "flow_int") <> 0 Then
	                Spec_Evaluate_DC_for_flow_loop = TheExec.Flow.var(temp_pin_info).value
	                outstring_evaluated = pininfo & ":" & condition_info & "=" & CStr(Spec_Evaluate_DC_for_flow_loop)
	                TheExec.Datalog.WriteComment "Calculate_Result:" & Trim(outstring_evaluated)
	            Else
	                Spec_Evaluate_DC_for_flow_loop = CStr(temp_pin_info)
	            End If
	        End If
	    End If
        Exit For
    Next site
        
''Debug.Print calc_info & "=" & Spec_Evaluate_DC_for_flow_loop

    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Spec_Evaluate_DC_for_flow_loop") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function Spec_Evaluate_AC(ByVal temp_pin_info As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim temp_pin_name, calc_info As String
    Dim temp_pininfo_arr() As String
    Dim i As Integer
    Dim fake_calc_info As String
    
    If InStr(temp_pin_info, "+") > 0 Or InStr(temp_pin_info, "-") > 0 Or InStr(temp_pin_info, "*") > 0 Or InStr(temp_pin_info, "/") > 0 Then
            calc_info = temp_pin_info
            
            temp_pin_info = Replace(temp_pin_info, "(", vbNullString)
            temp_pin_info = Replace(temp_pin_info, ")", vbNullString)
            temp_pin_info = Replace(temp_pin_info, "+", "~")
            temp_pin_info = Replace(temp_pin_info, "-", "~")
            temp_pin_info = Replace(temp_pin_info, "*", "~")
            temp_pin_info = Replace(temp_pin_info, "/", "~")
            
            fake_calc_info = Replace(calc_info, "(", " ( ")
            fake_calc_info = Replace(fake_calc_info, ")", " ) ")
            fake_calc_info = Replace(fake_calc_info, "+", " + ")
            fake_calc_info = Replace(fake_calc_info, "-", " - ")
            fake_calc_info = Replace(fake_calc_info, "*", " * ")
            fake_calc_info = Replace(fake_calc_info, "/", " / ")
            fake_calc_info = " " & fake_calc_info & " "
            
            temp_pininfo_arr = Split(temp_pin_info, "~")
            
            For i = 0 To UBound(temp_pininfo_arr)
                If InStr(temp_pininfo_arr(i), "_") <> 0 Then
                    temp_pin_name = temp_pininfo_arr(i)
                    temp_pininfo_arr(i) = CStr(TheExec.Specs.AC.item(mid(temp_pininfo_arr(i), 2)).ContextValue)
                    calc_info = Replace(fake_calc_info, " " & temp_pin_name & " ", temp_pininfo_arr(i), , 1)
                End If
            Next i
            Spec_Evaluate_AC = CStr(Evaluate(calc_info))
    Else
        If (InStr(temp_pin_info, "_") = 1) Then
            If TheExec.Specs.AC.Contains(mid(temp_pin_info, 2)) Then
                Spec_Evaluate_AC = CStr(TheExec.Specs.AC.item(mid(temp_pin_info, 2)).ContextValue)
            Else
                Spec_Evaluate_AC = CStr(temp_pin_info)
            End If
        Else
            Spec_Evaluate_AC = CStr(temp_pin_info)
        End If
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Spec_Evaluate_AC") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function SetupDCVI_ForceV(pin_name As String, arg_str As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    'PAD_MTR_ANALOG_TEST_P:SetupFV:Vprog,Irange,CustomizeWaitTime
    Dim arg_ary() As String
    Dim Vprog As Double, Iprog As Double, Irange As Double, CustomizeWaitTime As String
    Dim PowerType() As String
    Dim Factor As Long
    Dim NumTypes As Long
    Dim WaitTime As Double
    
    arg_ary = Split(arg_str, ",")
    Vprog = CDbl(Spec_Evaluate_DC(arg_ary(0)))
    Irange = CDbl(Spec_Evaluate_DC(arg_ary(1)))
    CustomizeWaitTime = arg_ary(2)
    
    Call TheExec.DataManager.GetChannelTypes(pin_name, NumTypes, PowerType())

    Select Case PowerType(0)
        Case "DCVI"
            Factor = 1
        Case "DCVIMerged"
            Factor = 2
        Case Else
    End Select
    
    If Irange > 2 * Factor Then
        Irange = 2 * Factor
        WaitTime = 1.6 * ms
    ElseIf Irange > 1 * Factor Then
        Irange = 2 * Factor
        WaitTime = 1.6 * ms
    ElseIf Irange > 0.2 * Factor Then
        Irange = 1 * Factor
        WaitTime = 1.6 * ms
    ElseIf Irange > 0.02 * Factor Then
        Irange = 0.2 * Factor
        WaitTime = 260 * us
    ElseIf Irange > 0.002 * Factor Then
        Irange = 0.02 * Factor
        WaitTime = 1.5 * ms
    ElseIf Irange > 0.0002 * Factor Then
        Irange = 0.002 * Factor
        WaitTime = 11 * ms
    ElseIf Irange > 0.00002 * Factor Then
        Irange = 0.0002 * Factor
        WaitTime = 1.4 * ms
    Else
        Irange = 0.00002 * Factor
        WaitTime = 6 * ms
    End If

    With TheHdw.DCVI.Pins(pin_name)
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = Vprog
        .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
        ''20161018 - Swap current and current range sequence to avoid mode alarm
''            .Current = Irange
''            .CurrentRange.Value = Irange
        .SetCurrentAndRange Irange, Irange
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With

    With TheHdw.DCVI.Pins(pin_name)
        .Meter.mode = tlDCVIMeterCurrent
        .Meter.CurrentRange.value = Irange
    End With
    If glb_Disable_CurrRangeSetting_Print = False Then
        TheExec.Datalog.WriteComment (TheExec.DataManager.instancename & " =====> Curr_meas Meter I range setting, " & pin_name & " =" & TheHdw.DCVI.Pins(pin_name).Meter.CurrentRange.value)
        TheExec.Datalog.WriteComment (TheExec.DataManager.instancename & " =====> Curr_meas Force Volt value, " & pin_name & " =" & Format(TheHdw.DCVI.Pins(pin_name).Voltage, "0.000"))
    End If
    If CustomizeWaitTime <> "" Then
        WaitTime = CDbl(CustomizeWaitTime)
    End If
    TheHdw.Wait (WaitTime)
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "SetupDCVI_ForceV") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function SetupDCVI_ForceI(pin_name As String, arg_str As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    'PAD_MTR_ANALOG_TEST_P:SetupFI:Vprog,Iprog
    Dim arg_ary() As String
    Dim Vprog As Double, Iprog As Double
    arg_ary = Split(arg_str, ",")
    Vprog = CDbl(Spec_Evaluate_DC(arg_ary(0)))
    Iprog = CDbl(Spec_Evaluate_DC(arg_ary(1)))
                
    With TheHdw.DCVI.Pins(pin_name) '' High impedence mode
        If Iprog = 0 Then
            '' 20150612 - High impedence mode
            ' Only required if force was previously connected
            .Disconnect tlDCVIConnectDefault
            ' Program the DCVI mapped to MyPin to high impedance mode
            .mode = tlDCVIModeHighImpedance
            ' Connect only the sense to use with high impedance mode
            .Connect tlDCVIConnectHighSense
            .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
            .Current = 0
        Else
            .mode = tlDCVIModeCurrent
            .Connect tlDCVIConnectDefault
            .Voltage = Vprog
            .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
            ''20170526-Add FI condition
            .CurrentRange.Autorange = True
            .Current = Iprog
        End If
        .VoltageRange.Autorange = True
'        thehdw.Wait (5 * ms)
        .Gate = True
    End With
    
     If glb_Disable_CurrRangeSetting_Print = False Then
        TheExec.Datalog.WriteComment (TheExec.DataManager.instancename & " =====> Volt_meas Force Current value, " & pin_name & " =" & Format(TheHdw.DCVI.Pins(pin_name).Current, "0.000000"))
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "SetupDCVI_ForceI") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MbistRetentionLevelWait_ForChar(mS_Time As Double, Retention_Voltage() As SiteDouble, Retention_Pins As PinList, RampStep As Double, Optional RampWaitTime As Double = 0.001)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
   ' thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered  'SEC DRAM
    
    'Dim Retention_Voltage As Double: Retention_Voltage = 0.5
    ''SWLINZA20171103, for ramp up/down retention voltage

    'Dim Retention_Pins As New PinList
    Dim Retention_Pins_Ary() As String
    Dim Retention_Pins_count As Long
    Dim RampDown_Time As Double: RampDown_Time = RampWaitTime 'RampDown_Time = 0
    Dim RampDown_Step As Double
    Dim Original_voltage() As New SiteDouble
    Dim DropVoltage() As New SiteDouble
    Dim DropVoltage_perStep() As New SiteDouble
    Dim Voltage_from_HW As String
    Dim ApplyPins As String
    Dim i, j As Integer
    Dim Flag_ApplyPower() As New SiteBoolean
    Dim site As Variant 'Carter, 20240304
    If RampStep = 0 Then
        RampDown_Step = 20 ' default RampDown_Step = 20
    Else
        RampDown_Step = RampStep
    End If
    
    TheExec.DataManager.DecomposePinList Retention_Pins, Retention_Pins_Ary(), Retention_Pins_count
    ReDim Original_voltage(Retention_Pins_count - 1) As New SiteDouble
    ReDim DropVoltage(Retention_Pins_count - 1) As New SiteDouble
    ReDim DropVoltage_perStep(Retention_Pins_count - 1) As New SiteDouble
    ReDim Flag_ApplyPower(Retention_Pins_count - 1) As New SiteBoolean
    
    TheExec.Datalog.WriteComment "********************************************************"
    TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
    
    For Each site In TheExec.sites
        For i = 0 To Retention_Pins_count - 1
            Flag_ApplyPower(i) = True
            Original_voltage(i) = Format(TheHdw.DCVS.Pins(Retention_Pins_Ary(i)).Voltage.Main, "0.000") 'Chris Modify
            'Chris add for offline simulation 20191023
            If TheExec.TesterMode = testModeOffline Then
                Original_voltage(i) = 5
            End If
            DropVoltage(i) = Original_voltage(i) - Retention_Voltage(i)
            DropVoltage_perStep(i) = Format((DropVoltage(i) / RampDown_Step), "0.000") 'Chris Modify
            If Format(DropVoltage(i), 3) = 0 Then Flag_ApplyPower(i) = False
        Next i
    
        '--------- Ramp down for retention voltage ------'
        For i = 0 To RampDown_Step - 1
            For j = 0 To Retention_Pins_count - 1
                If Flag_ApplyPower(j) = True Then
                    If i = RampDown_Step - 1 Then
                        TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main = Retention_Voltage(j)
                    Else
                        TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main = Original_voltage(j) - DropVoltage_perStep(j) * i
                    End If
                End If
            Next j
            TheHdw.Wait RampDown_Time / RampDown_Step  'remove extra wait time for each step
        Next i
    
        Voltage_from_HW = vbNullString
        ApplyPins = vbNullString
        Dim Current_PinCount As Long
        '--------- Read back retention voltage from HW ------'
        For j = 0 To Retention_Pins_count - 1
            If Flag_ApplyPower(j) = True Then
                If Current_PinCount = 0 Then
                    Voltage_from_HW = CStr(Format(TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main, "0.000")) '& " V" 'Chris Modify
                    ApplyPins = Retention_Pins_Ary(j)
                Else
                    Voltage_from_HW = Voltage_from_HW & "," & CStr(Format(TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main, "0.000")) '& " V" 'Chris Modify
                    ApplyPins = ApplyPins & "," & Retention_Pins_Ary(j)
                End If
                Current_PinCount = Current_PinCount + 1
            End If
        Next j
        TheExec.Datalog.WriteComment "--------- Site:" & site & "----------"
        TheExec.Datalog.WriteComment "*print: MbistRetention Pins " & ApplyPins
        TheExec.Datalog.WriteComment "*print: MbistRetention Volt " & Voltage_from_HW
    Next site
    
    '----- Retention Wait time 100 ms ------
    TheHdw.Wait mS_Time * 0.001
    'TheExec.Datalog.WriteComment "*************************************************"
    'TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
    'TheExec.Datalog.WriteComment "*print: MbistRetention Pins " & ApplyPins
    'TheExec.Datalog.WriteComment "*print: MbistRetention Volt " & Voltage_from_HW
'    TheExec.Datalog.WriteComment "*print: MbistRetention Voltage " & Retintion_voltage & " V*"
    TheExec.Datalog.WriteComment "********************************************************"
    DebugPrintFunc vbNullString
''
    '--------- Ramp up for retention voltage ------'
    For Each site In TheExec.sites
        For i = 0 To RampDown_Step - 1
            For j = 0 To Retention_Pins_count - 1
                If Flag_ApplyPower(j) = True Then
                    If i = RampDown_Step - 1 Then
                        TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main = Original_voltage(j)
                    Else
                        TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main = Retention_Voltage(j) + DropVoltage_perStep(j) * i
                    End If
                End If
            Next j
            TheHdw.Wait RampDown_Time / RampDown_Step 'remove extra wait time for each step
        Next i
    Next site
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "MbistRetentionLevelWait_ForChar") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function



Public Function Get_Tname_FromFlowSheet(Flow_Instance_Tname As String, HIO_PinName_Updated As Boolean) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim Current_FlowSheet_Loc As String
    Dim Current_Instance_Tname As String
    Dim CurrentSheet As Worksheet:: Set CurrentSheet = Nothing
    
    HIO_PinName_Updated = False
    Set CurrentSheet = ActiveWorkbook.Sheets(TheExec.Flow.Raw.SheetInRun)
    CurrentSheet.Select
'    Current_FlowSheet_Loc = theexec.Flow.Raw.SheetInRun + ":" + CStr(theexec.Flow.Raw.GetCurrentLineNumber + 5)
'    Current_Instance_Tname = CurrentSheet.Cells(theexec.Flow.Raw.GetCurrentLineNumber + 5, 9)
    Current_Instance_Tname = TheExec.Datalog.GetTestName(tl_Flowsheet_Testname)
    If UCase(Current_Instance_Tname) = UCase(Flow_Instance_Tname) Then Current_Instance_Tname = vbNullString
    If Current_Instance_Tname <> "" Then
'        Get_Tname_FromFlowSheet = Current_Instance_Tname
'        Flow_Instance_Tname = Current_Instance_Tname
        
        If UCase(Current_Instance_Tname) Like "*HAC*" Then Exit Function
        If UCase(Current_Instance_Tname) Like "*HIO*" Then
            Dim TNameSeg() As String
            ReDim TNameSeg(9) As String
            Dim SetupName As String
            Dim X_ApplyToPin As String
            Dim Y_ApplyToPin As String
            
            TNameSeg = Split(Current_Instance_Tname, "_")
            SetupName = TheExec.DevChar.Setups.ActiveSetupName
            With TheExec.DevChar.Setups(SetupName)
                If .Shmoo.axes.Count > 1 Then
                    X_ApplyToPin = .Shmoo.axes.item(tlDevCharShmooAxis_X).ApplyTo.Pins
                    Y_ApplyToPin = .Shmoo.axes.item(tlDevCharShmooAxis_Y).ApplyTo.Pins
                    TNameSeg(5) = Replace(X_ApplyToPin & "&" & Y_ApplyToPin, ",", "&")
'                    TNameSeg(5) = Replace(TNameSeg(5), ",", "&")
                Else
                    X_ApplyToPin = .Shmoo.axes.item(tlDevCharShmooAxis_X).ApplyTo.Pins
                    TNameSeg(5) = Replace(X_ApplyToPin, ",", "&")
                End If
            End With
            Flow_Instance_Tname = Merge_TName(TNameSeg)
            HIO_PinName_Updated = True
        Else
            Flow_Instance_Tname = Current_Instance_Tname
'            HIO_PinName_Updated = False
        End If
    End If
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Get_Tname_FromFlowSheet") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function



Public Function CZ_Shmoo_Info(ByRef SetupName As String, ByRef X_StepName As String, ByRef X_ApplyToPin As String, ByRef X_RangeFrom As Double, ByRef X_RangeTo As Double, ByRef x_stepsize As Double, _
                            Optional ByRef Y_StepName As String, Optional ByRef Y_ApplyToPin As String, Optional ByRef Y_RangeFrom As Double, Optional ByRef Y_RangeTo As Double, Optional ByRef y_stepsize As Double) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    If TheExec.DevChar.Setups.IsRunning = True Then
        SetupName = TheExec.DevChar.Setups.ActiveSetupName
        With TheExec.DevChar.Setups(SetupName)
            If .Shmoo.axes.Count > 1 Then
                X_StepName = .Shmoo.axes(tlDevCharShmooAxis_X).StepName
                Y_StepName = .Shmoo.axes(tlDevCharShmooAxis_Y).StepName
                X_ApplyToPin = .Shmoo.axes.item(tlDevCharShmooAxis_X).ApplyTo.Pins
                Y_ApplyToPin = .Shmoo.axes.item(tlDevCharShmooAxis_Y).ApplyTo.Pins
                X_RangeFrom = .Shmoo.axes(tlDevCharShmooAxis_X).Parameter.range.from
                Y_RangeFrom = .Shmoo.axes(tlDevCharShmooAxis_Y).Parameter.range.from
                X_RangeTo = .Shmoo.axes(tlDevCharShmooAxis_X).Parameter.range.to
                Y_RangeTo = .Shmoo.axes(tlDevCharShmooAxis_Y).Parameter.range.to
                x_stepsize = .Shmoo.axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
                y_stepsize = .Shmoo.axes(tlDevCharShmooAxis_Y).Parameter.range.StepSize
            Else
                X_StepName = .Shmoo.axes(tlDevCharShmooAxis_X).StepName
                X_ApplyToPin = .Shmoo.axes.item(tlDevCharShmooAxis_X).ApplyTo.Pins
                X_RangeFrom = .Shmoo.axes(tlDevCharShmooAxis_X).Parameter.range.from
                X_RangeTo = .Shmoo.axes(tlDevCharShmooAxis_X).Parameter.range.to
                x_stepsize = .Shmoo.axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
            End If
        End With
    End If

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "CZ_Shmoo_Info") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function CZ_TNum_Increment(Optional Flow_TestNumber As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
        
    Dim site As Variant
    
    If TheExec.DevChar.Setups.IsRunning = True Then
        For Each site In TheExec.sites.Active       '20190410 TER
              TheExec.Flow.TestNumber = TheExec.sites.item(site).TestNumber + 1
            Exit For
      Next site
    End If
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "CZ_TNum_Increment") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function



Public Function CZ_TNum_Decrement(Optional Flow_TestNumber As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
        
    Dim site As Variant
    
    If TheExec.DevChar.Setups.IsRunning = True Then
        For Each site In TheExec.sites.Active       '20190410 TER
            TheExec.Flow.TestNumber = TheExec.sites.item(site).TestNumber - 1
            Exit For
        Next site

    End If
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "CZ_TNum_Decrement") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function Spec_Evaluate_DC(ByVal temp_pin_info As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim temp_pin_name, calc_info As String
    Dim temp_pininfo_arr() As String
    Dim i As Integer
    Dim fake_calc_info As String
    Dim site As Variant 'Carter, 20240304
    
    If InStr(temp_pin_info, "+") > 0 Or InStr(temp_pin_info, "-") > 0 Or InStr(temp_pin_info, "*") > 0 Or InStr(temp_pin_info, "/") > 0 Then
        calc_info = temp_pin_info
        
        temp_pin_info = Replace(temp_pin_info, "(", vbNullString)
        temp_pin_info = Replace(temp_pin_info, ")", vbNullString)
        temp_pin_info = Replace(temp_pin_info, "+", "~")
        temp_pin_info = Replace(temp_pin_info, "-", "~")
        temp_pin_info = Replace(temp_pin_info, "*", "~")
        temp_pin_info = Replace(temp_pin_info, "/", "~")
        
        fake_calc_info = Replace(calc_info, "(", " ( ")
        fake_calc_info = Replace(fake_calc_info, ")", " ) ")
        fake_calc_info = Replace(fake_calc_info, "+", " + ")
        fake_calc_info = Replace(fake_calc_info, "-", " - ")
        fake_calc_info = Replace(fake_calc_info, "*", " * ")
        fake_calc_info = Replace(fake_calc_info, "/", " / ")
        fake_calc_info = " " & fake_calc_info & " "
        
        temp_pininfo_arr = Split(temp_pin_info, "~")
        
        For i = 0 To UBound(temp_pininfo_arr)
            If InStr(temp_pininfo_arr(i), "_") <> 0 Then
                temp_pin_name = temp_pininfo_arr(i)
                For Each site In TheExec.sites.Active
                    temp_pininfo_arr(i) = CStr(TheExec.Specs.DC.item(mid(temp_pininfo_arr(i), 2)).CurrentValue(site))
                    Exit For
                Next site
                calc_info = Replace(fake_calc_info, " " & temp_pin_name & " ", temp_pininfo_arr(i), , 1)
            End If
        Next i
        Spec_Evaluate_DC = CStr(Evaluate(calc_info))
    Else
        If (InStr(temp_pin_info, "_") = 1) Then
            If TheExec.Specs.DC.Contains(mid(temp_pin_info, 2)) Then
                For Each site In TheExec.sites.Active
                    Spec_Evaluate_DC = CStr(TheExec.Specs.DC.item(mid(temp_pin_info, 2)).CurrentValue(site))
                    Exit For
                Next site
            Else
                Spec_Evaluate_DC = CStr(temp_pin_info)
            End If
        Else
            Spec_Evaluate_DC = CStr(temp_pin_info)
        End If
    End If
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Spec_Evaluate_DC") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function SetPowerValue_Valt(ByVal pin As String, ByVal Pin_value As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim PinList() As String
    Dim PinNum As Integer
    Dim get_type As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var As Variant
    Dim PinName As String
    
    
    Call TheExec.DataManager.DecomposePinList(pin, PinList, typesCount)
    
    
    For Each var In PinList

        Call TheExec.DataManager.GetChannelTypes(var, typesCount, stringTypes)
        If (stringTypes(0) Like "DCVS*") Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = var + ":" + "VALT:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Alt.value), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "VALT:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Alt.value), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = var + ":" + "VALT:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Alt.value), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "VALT:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Alt.value), "0.000")
                End If
            Else
            'Do nothing
            End If
'            If g_Vbump_function = True Then 'add for SelSram function
'                 If Not g_CharInputString_Voltage_Dict.Exists(UCase(var)) = True Then
'                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(pin_value)
'                 Else
'                    g_CharInputString_Voltage_Dict.Remove (UCase(var))
'                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(pin_value)
'                 End If
'            Else
'                ''20200903 Solve Bora dsgrings alarm issue
'                ''1. while store interpose PrePat, voltage output selector would be Vmain, save value and pin in Valt
'                ''2. while restore original voltage, voltage output selector would be Valt, save value and pin in Vmain
'                If TheExec.DataManager.instancename Like "DSGRINGS_RNGV*" And Not (f_restorePrePat) Then
'                    TheHdw.DCVS.Pins(var).Voltage.Alt.value = CDbl(pin_value)
'                Else
'                    TheHdw.DCVS.Pins(var).Voltage.Main.value = CDbl(pin_value)     '' 20200903 Orignal Universal code
'                End If
                TheHdw.DCVS.Pins(var).Voltage.Alt.value = CDbl(Pin_value)
'            End If
            
        End If
        
cont:
    Next var
    
    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "SetPowerValue_Valt") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function Force_Condition_V(ByVal pin As String, ByVal force_val_Str As String, Optional restorepin As Boolean)
On Error GoTo errHandler
    ' Support ForceV condition of DCVI and DCVS in UF/UFP
    ' Replace ":setupfv:" and ":V:" condition
    ' Use Format--> Pin_Name:FV:ForceV_Val[@Ramping@Step@[Inverval_Time]],[I_clamp_Range],[CustomizeWaitTime]
    ' Ex. MTR_TD_B:FV:0.1@Ramping@10@0.001,0.02,0.001
    '======= Instrument list =======
    '=== I/O ===
    'HSDP     -> UP2200
    'HSD-U    -> UP1600
    '=== DCVS ===
    'VSM      -> VSM
    'HexVS    -> HexVS
    'VHDVS    -> UVS256
    'VS-800mA -> UVS256-HP
    'VS-5A    -> UVS64
    '=== DCVI ===
    'DC-07    -> UVI80
    'DC-30    -> DC30
    'DC-75    -> DC75
    
    '---- Parsing Force value condition ----  20230508
    Dim pin_condition_ary() As String
    Dim Pin_value As String
    Dim Ramp_Ary() As String
    Dim Ramp_Start As Double
    Dim Ramp_Stop As Double
    Dim Ramp_Step As Double
    Dim Ramp_StepVal As Double
    Dim Ramp_Delay As Double
    Dim Ramp_pin_value As Double
    
    
    'Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(pin, "V", force_val_Str))  'alen modify
    
    
    ' Use Format--> Pin_Name:FV:ForceV_Val[@Ramping@Step@[Inverval_Time]],[I_clamp_Range],[CustomizeWaitTime]
    ' Ex. MTR_TD_B:FV:0.1@Ramping@10@0.001,0.02,0.001
    pin_condition_ary = Split(force_val_Str, ",")
    
    Dim p_ary() As String, p_cnt As Long
    Dim PinList() As String
    Dim PinNum As Integer
    Dim get_type As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var As Variant
    Dim PinName As String
    Dim ClampI_Range As Double
    Dim site As Variant 'Carter, 20240304
    Dim pin_instrument_type as string
	
	pin_instrument_type = UCase(SortPinInstrument(pin))

    If InStr(force_val_Str, "[") > 0 Then ''e.g. PinA:V:0.5*[SrcCodeIndx]+0.001*[SrcCodeIndx1]
        Dim FlowVarStr As String
        Dim TempIdx As Long
        Dim SplitLeftBigColon() As String
        SplitLeftBigColon = Split(force_val_Str, "[")
        For TempIdx = 0 To UBound(SplitLeftBigColon)
            If InStr(SplitLeftBigColon(TempIdx), "]") > 0 Then
                FlowVarStr = Split(SplitLeftBigColon(TempIdx), "]")(0)
                force_val_Str = Replace(force_val_Str, "[" & FlowVarStr & "]", CStr(TheExec.Flow.var(FlowVarStr).value))
            End If
        Next TempIdx
    End If
    
    If pin_instrument_type = "HSD-U" Or pin_instrument_type = "HSDP" Then
        '=============== UP1600/UP2200 ==================
        TheHdw.Digital.Pins(pin).Disconnect
        
        ClampI_Range = 0#
        If UBound(pin_condition_ary) >= 1 Then
            If pin_condition_ary(1) <> "" Then
                ClampI_Range = CDbl(pin_condition_ary(1))
            Else
                ClampI_Range = 0.02  'Setup Clamp I range
            End If
        Else
            ClampI_Range = 0.02  'Setup Clamp I range
        End If
        
        'Do not use the 50 mA current range.
        'These digital instruments do not allow a simultaneous connection of both PPMU and PE functional
        'relays due to resource sharing in the 50 mA current range. This method can connect these relays at the same time,
        'unless you try to do the following:
        'Connect both PPMU and PE functional relays while in the 50 mA current range
        'Set the 50 mA current range while both PPMU and PE functional relays are connected
        
        If InStr(UCase(force_val_Str), "RAMPING") <> 0 Then
            '==== Add Ramping voltage =====
            'MTR_TD_B:FV:0.1@Ramping@10@0.001,0.02,0.001
            '---- Parsing Force value condition ----  20230508
            Ramp_Ary = Split(pin_condition_ary(0), "@")
            For Each var In PinList
                Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(var, "V", Ramp_Ary(0)))
                Ramp_Start = TheHdw.PPMU.Pins(var).Voltage.value
                Ramp_Stop = CDbl(Pin_value)
                'Ramp_Stop = CDbl(Ramp_Ary(0))
                Ramp_Step = CDbl(Ramp_Ary(2))
                Ramp_Delay = CDbl(Ramp_Ary(3))
                Ramp_StepVal = (Ramp_Stop - Ramp_Start) / Ramp_Step   'Per step +/- voltage value
                
                If (PrePatStore = True) Then
                    If (PrePat_Restore_String = "") Then
                        PrePat_Restore_String = var + ":V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" & CStr(Ramp_Delay) + "," & CStr(ClampI_Range)
                    Else
                        PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range)
                    End If
                ElseIf (PreMeasStore = True) Then
                    If (PreMeas_Restore_String = "") Then
                        PreMeas_Restore_String = var + ":V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," & CStr(ClampI_Range)
                    Else
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range)
                    End If
                End If
                
                If (PrePatStore = True) Then
                    If (PrePat_Restore_String = "") Then
                        PrePat_Restore_String = var + ":DisConnectPPMU;" + var + ":ConnectDigital"
                    Else
                        PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":DisConnectPPMU;" + var + ":ConnectDigital"
                    End If
                ElseIf (PreMeasStore = True) Then
                    If (PreMeas_Restore_String = "") Then
                        PreMeas_Restore_String = var + ":DisConnectPPMU;" + var + ":ConnectDigital"
                    Else
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":DisConnectPPMU;" + var + ":ConnectDigital"
                    End If
                End If
                
                For Ramp_pin_value = Ramp_Start To Ramp_Stop Step Ramp_StepVal
                    If Ramp_pin_value = Ramp_Start Then
                        TheHdw.PPMU.Pins(var).ForceV (CDbl(Ramp_pin_value)), 0.02
                        TheHdw.PPMU.Pins(var).Connect
                        TheHdw.PPMU.Pins(var).Gate = tlOn
                    Else
                        TheHdw.PPMU.Pins(var).ForceV (CDbl(Ramp_pin_value)), ClampI_Range
                    End If
                    TheHdw.Wait (Ramp_Delay)    'Unit : sec
                    TheExec.Datalog.WriteComment "Force Voltage[Rampping]: Pin: " & var & ", Voltage: " & CStr(Format(CDbl(Ramp_pin_value), "0.0000"))
                Next Ramp_pin_value
            Next var
        Else
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(pin).Voltage, "0.000000")) & "," & CStr(ClampI_Range)
                Else
                    PrePat_Restore_String = PrePat_Restore_String & ";" & pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(pin).Voltage, "0.000000")) & "," & CStr(ClampI_Range)
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(pin).Voltage, "0.000000")) & "," & CStr(ClampI_Range)
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String & ";" & pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(pin).Voltage, "0.000000")) & "," & CStr(ClampI_Range)
                End If
            End If
            
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = pin + ":DisConnectPPMU;" + pin + ":ConnectDigital"
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + pin + ":DisConnectPPMU;" + pin + ":ConnectDigital"
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = pin + ":DisConnectPPMU;" + pin + ":ConnectDigital"
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + pin + ":DisConnectPPMU;" + pin + ":ConnectDigital"
                End If
            End If
        
            Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(pin, "V", pin_condition_ary(0)))
            'TheHdw.PPMU.Pins(pin).ForceV (CDbl(pin_condition_ary(0))), ClampI_Range
			
            If Pin_value = "-999" Then
                '[20231107][All][Neil] Apply Global value [SHMOO_GLB] for PPMU pin sweep voltage when Shmoo processing
                For Each site In TheExec.sites
                    TheHdw.PPMU.Pins(pin).ForceV SD_Shmoo_GLB_Val, 0.02
                Next site
            Else
                TheHdw.PPMU.Pins(pin).ForceV (CDbl(Pin_value)), ClampI_Range
            End If			
			
            TheHdw.PPMU.Pins(pin).Connect
            TheHdw.PPMU.Pins(pin).Gate = tlOn
			
        End If
                
        '==== Set CustomizeWaitTime =====
        If UBound(pin_condition_ary) = 2 Then
            TheHdw.Wait (pin_condition_ary(2))
        End If
    ElseIf pin_instrument_type = "DC-07" Or pin_instrument_type = "DC-30" Or pin_instrument_type = "DC-75" Then    'DCVI
        Dim PowerType() As String
        Dim NumTypes As Long
        Dim WaitTime As Double
        Dim Factor As Long
        Dim Irange As Double
        
        ClampI_Range = 0#
        'MTR_TD_B:FV:0.1@Ramping@10@0.001,0.02,0.001
        '==== Set I range ==== 'nn_test_0508
        If UBound(pin_condition_ary) >= 1 Then
            If pin_condition_ary(1) <> "" Then
                ClampI_Range = pin_condition_ary(1)   'Get Setup/Meas Irange from input parameter
            Else
                ClampI_Range = pc_Def_UVI80_Init_MeasCurrRange '20uA
            End If
        Else
            ClampI_Range = pc_Def_UVI80_Init_MeasCurrRange '20uA
        End If
        
        For Each var In PinList
            Call TheExec.DataManager.GetChannelTypes(var, NumTypes, PowerType())
                
            If restorepin = True And LCase(TheExec.DataManager.PinType(LCase(var))) Like "analog" = True Then
                With TheHdw.DCVI.Pins(var)
                    .Gate = False
                    .Disconnect
                End With
            Else
                Select Case PowerType(0)
                    Case "DCVI"
                        Factor = 1
                    Case "DCVIMerged"
                        Factor = 2
                    Case Else
                End Select
            
                If Irange > 2 * Factor Then
                    Irange = 2 * Factor
                    WaitTime = 1.6 * ms
                ElseIf Irange > 1 * Factor Then
                    Irange = 2 * Factor
                    WaitTime = 1.6 * ms
                ElseIf Irange > 0.2 * Factor Then
                    Irange = 1 * Factor
                    WaitTime = 1.6 * ms
                ElseIf Irange > 0.02 * Factor Then
                    Irange = 0.2 * Factor
                    WaitTime = 260 * us
                ElseIf Irange > 0.002 * Factor Then
                    Irange = 0.02 * Factor
                    WaitTime = 1.5 * ms
                ElseIf Irange > 0.0002 * Factor Then
                    Irange = 0.002 * Factor
                    WaitTime = 11 * ms
                ElseIf Irange > 0.00002 * Factor Then
                    Irange = 0.0002 * Factor
                    WaitTime = 1.4 * ms
                Else
                    Irange = 0.00002 * Factor
                    WaitTime = 6 * ms
                End If
    
                With TheHdw.DCVI.Pins(var)
                    .Gate = False
                    .mode = tlDCVIModeVoltage
                    .Voltage = 0#
                    .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
                    ''20161018 - Swap current and current range sequence to avoid mode alarm
                    .SetCurrentAndRange ClampI_Range, ClampI_Range
                    .Connect tlDCVIConnectDefault
                    .Gate = True
                End With
        
                With TheHdw.DCVI.Pins(var)
                    .Meter.mode = tlDCVIMeterCurrent
                    .Meter.CurrentRange.value = ClampI_Range
                End With
                
                If InStr(UCase(force_val_Str), "RAMPING") <> 0 Then
                    '==== Add Ramping voltage =====
                    'MTR_TD_B:FV:0.1@Ramping@10@0.001,0.02,0.001
                    '---- Parsing Force value condition ----  20230508
                    Ramp_Ary = Split(pin_condition_ary(0), "@")
                    Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(var, "V", Ramp_Ary(0)))
                    
                    Ramp_Start = TheHdw.DCVI.Pins(var).Voltage.value
                    Ramp_Stop = CDbl(Pin_value)
                    'Ramp_Stop = CDbl(Ramp_Ary(0))
                    Ramp_Step = CDbl(Ramp_Ary(2))
                    Ramp_Delay = CDbl(Ramp_Ary(3))
                    Ramp_StepVal = (Ramp_Stop - Ramp_Start) / Ramp_Step   'Per step +/- voltage value
                    
                    If (PrePatStore = True) Then
                        If (PrePat_Restore_String = "") Then
                            PrePat_Restore_String = var + ":V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range) + ";" + var + ":" + "restoredcvi"
                        Else
                            PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range) + ";" + var + ":" + "restoredcvi"
                        End If
                    ElseIf (PreMeasStore = True) Then
                        If (PreMeas_Restore_String = "") Then
                            PreMeas_Restore_String = var + ":V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range) + ";" + var + ":" + "restoredcvi"
                        Else
                            PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range) + ";" + var + ":" + "restoredcvi"
                        End If
                    End If
                
                    For Ramp_pin_value = Ramp_Start To Ramp_Stop Step Ramp_StepVal
                        TheHdw.DCVI.Pins(var).Voltage = CDbl(Ramp_pin_value)
                        TheHdw.Wait (Ramp_Delay)    'Unit : sec
                        TheExec.Datalog.WriteComment "Force Voltage[Rampping]: Pin: " & var & ", Voltage: " & CStr(Format(CDbl(Ramp_pin_value), "0.0000"))
                    Next Ramp_pin_value
                Else
                    If (PrePatStore = True) Then
                        If (PrePat_Restore_String = "") Then
                            PrePat_Restore_String = var + ":V:" + CStr(TheHdw.DCVI.Pins(var).Voltage.value) + ";" + var + ":" + "restoredcvi"
                        Else
                            PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":V:" + CStr(TheHdw.DCVI.Pins(var).Voltage.value) + ";" + var + ":" + "restoredcvi"
                        End If
                    ElseIf (PreMeasStore = True) Then
                        If (PreMeas_Restore_String = "") Then
                            PreMeas_Restore_String = var + ":V:" + CStr(TheHdw.DCVI.Pins(var).Voltage.value) + ";" + var + ":" + "restoredcvi"
                        Else
                            PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":V:" + CStr(TheHdw.DCVI.Pins(var).Voltage.value) + ";" + var + ":" + "restoredcvi"
                        End If
                    End If
                    Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(var, "V", pin_condition_ary(0)))
                    TheHdw.DCVI.Pins(var).Voltage = CDbl(Pin_value)
                End If
            End If
        Next var
            
        '==== Set CustomizeWaitTime =====
        If UBound(pin_condition_ary) = 2 Then
            TheHdw.Wait (pin_condition_ary(2))
        Else
            TheHdw.Wait (WaitTime)
        End If

    ElseIf pin_instrument_type = "HEXVS" Or pin_instrument_type = "VHDVS" Or pin_instrument_type = "VS-800MA" Or pin_instrument_type = "VS-5A" Then
        '=============== DCVS==================
        'Stop
        ClampI_Range = 0#
        'MTR_TD_B:FV:0.1@Ramping@10@0.001,0.02,0.001
        '==== Set I range ==== 'nn_test_0508
        If UBound(pin_condition_ary) >= 1 Then
            If pin_condition_ary(1) <> "" Then
                ClampI_Range = CDbl(pin_condition_ary(1))
            Else
            If pin_instrument_type = "HEXVS" Then ClampI_Range = 0.2                                    'HexVs
            If pin_instrument_type = "VHDVS" Then ClampI_Range = 0.2                                    'UVS256
            If pin_instrument_type = "VS-800MA" Then ClampI_Range = pc_Def_UVS256HP_Init_MeasCurrRange  'UVS256-HP
            If pin_instrument_type = "VS-5A" Then ClampI_Range = pc_Def_UVS64_Init_MeasCurrRange        'UVS64
            End If
        Else
            If pin_instrument_type = "HEXVS" Then ClampI_Range = 0.2                                    'HexVs
            If pin_instrument_type = "VHDVS" Then ClampI_Range = 0.2                                    'UVS256
            If pin_instrument_type = "VS-800MA" Then ClampI_Range = pc_Def_UVS256HP_Init_MeasCurrRange  'UVS256-HP
            If pin_instrument_type = "VS-5A" Then ClampI_Range = pc_Def_UVS64_Init_MeasCurrRange        'UVS64
        End If
        Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(pin, "V", force_val_Str))  'avoid Calculate_Result print two times
        For Each var In PinList
            If g_Vbump_function = True Then 'add for SelSram function           'Need to confirm -- 20230418
                 If Not g_CharInputString_Voltage_Dict.Exists(UCase(var)) = True Then
                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(Pin_value)
                 Else
                    g_CharInputString_Voltage_Dict.Remove (UCase(var))
                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(Pin_value)
                 End If
            Else
                Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(var, "V", pin_condition_ary(0)))

                If InStr(UCase(force_val_Str), "RAMPING") <> 0 Then
                    '==== Add Ramping voltage =====
                    'MTR_TD_B:FV:0.1@Ramping@10@0.001,0.02,0.001
                    '---- Parsing Force value condition ----  20230508
                    Ramp_Ary = Split(pin_condition_ary(0), "@")
                    Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(var, "V", Ramp_Ary(0)))

                    Ramp_Start = TheHdw.DCVS.Pins(var).Voltage.Main.value
                    Ramp_Stop = CDbl(Pin_value)
                    'Ramp_Stop = CDbl(Ramp_Ary(0))
                    Ramp_Step = CDbl(Ramp_Ary(2))
                    Ramp_Delay = CDbl(Ramp_Ary(3))
                    Ramp_StepVal = (Ramp_Stop - Ramp_Start) / Ramp_Step   'Per step +/- voltage value
                    
                    If (PrePatStore = True) Then
                        If (PrePat_Restore_String = "") Then
                            PrePat_Restore_String = var + ":" + "V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range)
                        Else
                            PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range)
                        End If
                        If TheHdw.DCVS.Pins(var).Gate = False Then PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "disconnectdcvs"
                    ElseIf (PreMeasStore = True) Then
                        If (PreMeas_Restore_String = "") Then
                            PreMeas_Restore_String = var + ":" + "V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range)
                        Else
                            PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "V:" + CStr(Ramp_Start) + "@Ramping@" + CStr(Ramp_Step) + "@" + CStr(Ramp_Delay) + "," + CStr(ClampI_Range)
                        End If
                        If TheHdw.DCVS.Pins(var).Gate = False Then PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "disconnectdcvs"
                    End If
                    
                    If TheHdw.DCVS.Pins(var).Gate = False Then
                        With TheHdw.DCVS.Pins(var)
							.Voltage.Output = tlDCVSVoltageMain
                            .Disconnect tlDCVSConnectDefault
                            .Meter.mode = tlDCVSMeterCurrent
                            .mode = tlDCVSModeVoltage
                            .SetCurrentRanges CDbl(ClampI_Range), CDbl(ClampI_Range)
                            .Voltage.value = 0#
                            .Connect tlDCVSConnectDefault
                            .Gate = True
                        End With
                    Else
                    End If
                    For Ramp_pin_value = Ramp_Start To Ramp_Stop Step Ramp_StepVal
                        TheHdw.DCVS.Pins(var).Voltage.Main.value = CDbl(Ramp_pin_value)
                        TheHdw.Wait (Ramp_Delay)    'Unit : sec
                        TheExec.Datalog.WriteComment "Force Voltage[Rampping]: Pin: " & pin & ", Voltage: " & CStr(Format(CDbl(Ramp_pin_value), "0.0000"))
                    Next Ramp_pin_value
                Else
                    If (PrePatStore = True) Then
                        If (PrePat_Restore_String = "") Then
                            PrePat_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.000")
                        Else
                            PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.000")
                        End If
                        If TheHdw.DCVS.Pins(var).Gate = False Then PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "disconnectdcvs"
                    ElseIf (PreMeasStore = True) Then
                        If (PreMeas_Restore_String = "") Then
                            PreMeas_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.000")
                        Else
                            PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.value), "0.000")
                        End If
                        If TheHdw.DCVS.Pins(var).Gate = False Then PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "disconnectdcvs"
                    End If
                
                    If TheHdw.DCVS.Pins(var).Gate = False Then
                        With TheHdw.DCVS.Pins(var)
							.Voltage.Output = tlDCVSVoltageMain
                            .Disconnect tlDCVSConnectDefault
                            .Meter.mode = tlDCVSMeterCurrent
                            .mode = tlDCVSModeVoltage
                            .SetCurrentRanges CDbl(ClampI_Range), CDbl(ClampI_Range)
                            .Voltage.value = 0#
                            .Connect tlDCVSConnectDefault
                            .Gate = True
                        End With
                    Else
                    End If
                    Pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(var, "V", pin_condition_ary(0)))
                    TheHdw.DCVS.Pins(var).Voltage.Main.value = CDbl(Pin_value)
                End If
            End If
            
            If restorepin = True And LCase(TheExec.DataManager.PinType(LCase(var))) Like "analog" = True Then
                With TheHdw.DCVS.Pins(var)
                    .Gate = False
                    .Disconnect
                End With
            Else
            
            End If
        Next var
        '==== Set CustomizeWaitTime =====
        If UBound(pin_condition_ary) = 2 Then
            TheHdw.Wait (pin_condition_ary(2))
        End If
    Else
        Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Not found suiteble Instrument type!!!")
    End If
    
    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Force_Condition_V") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

'[20231221][All] New SetForceCondition method
Public Function Force_Condition_I(ByVal pin As String, ByVal force_val_Str As String)
On Error GoTo errHandler
    ' Support ForceI condition of DCVI and DCVS in UF/UFP
    ' Replace ":setupfi:" and ":I:" condition
    ' Use Format--> PIN_NAME:FI:IProg,[VProg]
    ' Ex. MTR_TD_B:FI:0.1,0.001
    '======= Instrument list =======
    '=== I/O ===
    'HSDP     -> UP2200
    'HSD-U    -> UP1600
    '=== DCVS ===
    'VSM      -> VSM
    'HexVS    -> HexVS
    'VHDVS    -> UVS256
    'VS-800mA -> UVS256-HP
    'VS-5A    -> UVS64
    '=== DCVI ===
    'DC-07    -> UVI80
    'DC-30    -> DC30
    'DC-75    -> DC75
    
    '---- Parsing Force value condition ----  20230508
    Dim pin_condition_ary() As String
    Dim Pin_value As String
    Dim Ramp_Ary() As String
    Dim Ramp_Start As Double
    Dim Ramp_Stop As Double
    Dim Ramp_Step As Double
    Dim Ramp_StepVal As Double
    Dim Ramp_Delay As Double
    Dim Ramp_pin_value As Double
    
    'Use for PPMU pin setting FV(Iclamp)
    'Format Pin_Name:FI:ForceI_Val,V_clamp_Range
    'Format Pin_Name:FI:ForceI_Val,V_clampHi&V_clampLo
    'Ex. MTR_TD_B:FI:0.1,0.02,0.001
    pin_condition_ary = Split(force_val_Str, ",")
    
    Dim p_ary() As String, p_cnt As Long
    Dim PinList() As String
    Dim PinNum As Integer
    Dim get_type As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var As Variant
    Dim PinName As String
    
    Dim V_ClampAry() As String
    Dim V_ClampLo As Double
    Dim V_ClampHi As Double
    Dim leader_pin As String
    Dim PinArray() As String
    
    Dim pin_instrument_type as string
	
	pin_instrument_type = UCase(SortPinInstrument(pin))
    
    If pin_instrument_type = "HSD-U" Or pin_instrument_type = "HSDP" Then
        '=============== UP2200/UP1600 ==================
        TheHdw.Digital.Pins(pin).Disconnect
'        V_ClampLo = CStr(Format(TheHdw.PPMU.Pins(pin).ClampVLo, "0.0000"))
'        V_ClampHi = CStr(Format(TheHdw.PPMU.Pins(pin).ClampVHi, "0.0000"))
        
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(pin).Current, "0.000000"))
            Else
                PrePat_Restore_String = PrePat_Restore_String & ";" & pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(pin).Current, "0.000000"))
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(pin).Current, "0.000000"))
            Else
                PreMeas_Restore_String = PreMeas_Restore_String & ";" & pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(pin).Current, "0.000000"))
            End If
        End If
        
        '==== Setup Voltage Clamp =====
        If (UBound(pin_condition_ary)) >= 1 Then
            If pin_condition_ary(1) <> "" Then
                V_ClampAry = Split(pin_condition_ary(1), "&")
                'Format Pin_Name:FI:ForceI_Val,V_clampHi&V_clampLo
                V_ClampHi = CDbl(V_ClampAry(0))
                V_ClampLo = CDbl(V_ClampAry(1))
            Else
                V_ClampHi = CDbl(V_ClampAry(0))
                V_ClampLo = CDbl(V_ClampAry(1))
            End If
        End If
        
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = pin + ":DisConnectPPMU"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + pin + ":DisConnectPPMU"
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = pin + ":DisConnectPPMU"
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + pin + ":DisConnectPPMU"
            End If
        End If
'        For Each var In PinList
            With TheHdw.PPMU.Pins(pin)
                '.mode = tlPPMUForceIMeasureV
                .ForceI (pin_condition_ary(0))
'                .ClampVLo = V_ClampLo
'                .ClampVHi = V_ClampHi
                .Connect
                .Gate = tlOn
            End With
'        Next var
    ElseIf pin_instrument_type = "DC-07" Or pin_instrument_type = "DC-30" Or pin_instrument_type = "DC-75" Then    'DCVI
        '=============== DCVI==================
        '==== Setup Voltage Clamp =====
        V_ClampHi = TheHdw.DCVI.Pins(pin).VoltageRange.value
        
        TheHdw.DCVI.Pins(pin).mode = tlDCVIModeCurrent
        
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                'PrePat_Restore_String = pin + ":" + "I:" + Format(CStr(thehdw.DCVI.Pins(pin).Current), "0.0000") + "," + CStr(V_ClampHi)
                PrePat_Restore_String = pin + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(pin).Current), "0.0000")
            Else
                'PrePat_Restore_String = PrePat_Restore_String + ";" + pin + ":" + "I:" + Format(CStr(thehdw.DCVI.Pins(pin).Current), "0.0000") + "," + CStr(V_ClampHi)
                PrePat_Restore_String = PrePat_Restore_String + ";" + pin + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(pin).Current), "0.0000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                'PreMeas_Restore_String = pin + ":" + "I:" + Format(CStr(thehdw.DCVI.Pins(pin).Current), "0.0000") + "," + CStr(V_ClampHi)
                PreMeas_Restore_String = pin + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(pin).Current), "0.0000")
            Else
                'PreMeas_Restore_String = PreMeas_Restore_String + ";" + pin + ":" + "I:" + Format(CStr(thehdw.DCVI.Pins(pin).Current), "0.0000") + "," + CStr(V_ClampHi)
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + pin + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(pin).Current), "0.0000")
            End If
        End If
        
        '==== Setup Voltage Clamp =====
        If (UBound(pin_condition_ary)) >= 1 Then
            If pin_condition_ary(1) <> "" Then
                V_ClampHi = CDbl(pin_condition_ary(1))
            End If
        End If
        
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = pin + ":" + "restoredcvi"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + pin + ":" + "restoredcvi"
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = pin + ":" + "restoredcvi"
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + pin + ":" + "restoredcvi"
            End If
        End If
'        Call TheExec.DataManager.DecomposePinList(pin, PinList, typesCount)
        
'        For Each var In PinList
            With TheHdw.DCVI.Pins(pin) '' High impedence mode
                If CDbl(pin_condition_ary(0)) = 0 Then
                    '' 20150612 - High impedence mode
                    ' Only required if force was previously connected
                    .Disconnect tlDCVIConnectDefault
                    ' Program the DCVI mapped to MyPin to high impedance mode
                    .mode = tlDCVIModeHighImpedance
                    ' Connect only the sense to use with high impedance mode
                    .Connect tlDCVIConnectHighSense
                    .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
                    .Current = 0
                Else
                    .mode = tlDCVIModeCurrent
                    .Connect tlDCVIConnectDefault
                    .Voltage = V_ClampHi               'Voltage Clamp
                    .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
                    ''20170526-Add FI condition
                    .CurrentRange.Autorange = True
                    .Current = CDbl(pin_condition_ary(0))
                End If
                .VoltageRange.Autorange = True
                .Gate = True
            End With
'        Next var
        'Stop
    ElseIf pin_instrument_type = "VHDVS" Or pin_instrument_type = "VS-800MA" Or pin_instrument_type = "VS-5A" Then
        '=============== DCVS==================
        If UCase(gl_GetInstrument_Dic(LCase(leader_pin))) = "VHDVS" Then
            Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "UVS256 : No support force current mode!!!")
        Else
            '==== Setup Voltage Clamp =====
            If (PrePatStore = True) Then
                If (Format(CStr(TheHdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value), "0.000") = 0) Or TheHdw.DCVS.Pins(pin).Gate = False Then
                    PrePat_Restore_String = PrePat_Restore_String + ";" + pin + ":" + "disconnectdcvs"
                Else
                    If (PrePat_Restore_String = "") Then
                        'PrePat_Restore_String = pin + ":" + "I:" + Format(CStr(thehdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value), "0.000") + "," + CStr(V_ClampHi)
                        PrePat_Restore_String = pin + ":" + "I:" + Format(CStr(TheHdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value), "0.000")
                    Else
                        'PrePat_Restore_String = PrePat_Restore_String + ";" + pin + ":" + "I:" + Format(thehdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value, "0.000") + "," + CStr(V_ClampHi)
                        PrePat_Restore_String = PrePat_Restore_String + ";" + pin + ":" + "I:" + Format(TheHdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value, "0.000")
                    End If
                End If
            ElseIf (PreMeasStore = True) Then
                If (Format(CStr(TheHdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value), "0.000") = 0) Or TheHdw.DCVS.Pins(pin).Gate = False Then
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + pin + ":" + "disconnectdcvs"
                Else
                    If (PreMeas_Restore_String = "") Then
                        'PreMeas_Restore_String = pin + ":" + "I:" + Format(CStr(thehdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value), "0.000") + "," + CStr(V_ClampHi)
                        PreMeas_Restore_String = pin + ":" + "I:" + Format(CStr(TheHdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value), "0.000")
                    Else
                        'PreMeas_Restore_String = PreMeas_Restore_String + ";" + pin + ":" + "I:" + Format(CStr(thehdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value), "0.000") + "," + CStr(V_ClampHi)
                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + pin + ":" + "I:" + Format(CStr(TheHdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value), "0.000")
                    End If
                End If
            End If
            
            If (UBound(pin_condition_ary)) >= 1 Then
                If pin_condition_ary(1) <> "" Then
                    V_ClampHi = CDbl(pin_condition_ary(1))
                End If
            End If
'            Call TheExec.DataManager.DecomposePinList(pin, PinList, typesCount)
            
'            For Each var In PinList
                With TheHdw.DCVS.Pins(pin)
                    If (CDbl(pin_condition_ary(0)) = 0) And .Gate <> True Then
                        .Disconnect
                        .mode = tlDCVSModeHighImpedance
                        .Meter.mode = tlDCVSMeterVoltage
                        .Gate = True
                        .Connect
                    Else
                        .VoltageRange.value = VoltageRange_Select(CDbl(pin_condition_ary(0)))
                        
                        If CDbl(V_ClampHi) = 0 Then
                            If CDbl(pin_condition_ary(0)) > 0 Then
                                .Voltage.value = 1 + 0.5
                            Else
                                '---------------------------------------UFP_Corr fix 20200413
                                '.VoltageRange.value = 5.5
                                '---------------------------------------UFP_Corr fix 20200413
                                .Voltage.value = -1 - 0.5
                            End If
                        Else
                            If CDbl(pin_condition_ary(0)) > 0 Then
                                .Voltage.value = CDbl(V_ClampHi) + 0.5
                            Else
                                '---------------------------------------UFP_Corr fix 20200413
                                '.VoltageRange.value = 5.5
                                '---------------------------------------UFP_Corr fix 20200413
                                .Voltage.value = -CDbl(V_ClampHi) - 0.5
                            End If
                        End If
                        .CurrentRange.value = Abs(CDbl(pin_condition_ary(0)))
                        .CurrentLimit.Source.FoldLimit.level.value = Abs(CDbl(pin_condition_ary(0)))
                        .CurrentLimit.Sink.FoldLimit.level.value = Abs(CDbl(pin_condition_ary(0)))
                        .Meter.mode = tlDCVSMeterVoltage
                        .Gate = True
                        .Connect
                        '---------------------------------------UFP_Corr fix 20200413
                        .mode = tlDCVSModeCurrent
                        '---------------------------------------UFP_Corr fix 20200413
                    End If
                End With
'            Next var
        End If
    Else
        Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Not found suiteble Instrument type!!!")
    End If
    
    '==== Set CustomizeWaitTime =====
    If UBound(pin_condition_ary) = 2 Then
        TheHdw.Wait (pin_condition_ary(2))
    End If
    
    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_ForceCondition", "Force_Condition_I") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
