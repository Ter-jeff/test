Attribute VB_Name = "LIB_DC"
#Const isUFP = True
Option Explicit

Private Const LIB_DC = "LIB_DC"

Public Flag_DC_ParsingRecoveryTable As Boolean   '''20210831 add by kelly
Public Const MaxLoadCount = 4           ''20210811 add for MeasRecoveryTimes
Public Const MaxCcodeCount_64 = 21           ''20210831 add for cCode_TYPE
Public Const MaxCcodeCount_256 = 16           ''20210831 add for cCode_TYPE
Public DC_SettleTime_UVS256HP() As UVS256HP_TYPE ''20210811 add for MeasRecoveryTimes
Public DC_SettleTime_UVS64() As UVS64_TYPE ''20210811 add for MeasRecoveryTimes
Public Dibcap() As DIBCAP_TYPE ''20210823 add for dibcap_cp
Public C_code_UVS64() As cCode_TYPE
Public C_code_UVS256HP() As cCode_TYPE
Public Dict_Dibcap_PowerPinName As New Dictionary


'===========================================================================================================
' The data structure is for storing data of the MeasRecoveryTimes
'===========================================================================================================
'''20210817
Public Type UVS256HP_TYPE '''For read C load Table (MeasRecoveryTimes)
    BW_min As Long
    BW_max As Long
    dict_CurrentRange As New Dictionary  ''20210811 add for MeasRecoveryTimes
End Type

Public Type UVS64_TYPE '''For read C load Table (MeasRecoveryTimes)
    BW_min As Long
    BW_max As Long
    dict_CurrentRange As New Dictionary  ''20210811 add for MeasRecoveryTimes
End Type

Public Type DIBCAP_TYPE '''For dibcap_cp table
    powerPin As String
    Merge As Integer
    CapValue As Double
    offset As Double
    SingleEndCap As Double
End Type

'===========================================================================================================
' The data structure is for c_code calculate
'===========================================================================================================
'''20210831
Public Type cCode_TYPE
    min As Double
    max As Double
    C_code As Integer
End Type

' Enum Status Num
Public Enum gateStatus ' Gate Status
    gOn = 1
    gOff = 2
    gKeep = 3
End Enum

Public Enum ConnectedStatus ' Connect Status
    Connect = 1
    Disconnect = 2
    KeepStatus = 3
End Enum

Public Enum forceType ' Force type
    fVoltage = 1
    fCurrent = 2
    fNone = 3
End Enum

Public Enum DCVSoutput ' DCVS output status
    dKeep = 1
    dVmain = 2
    dVAlt = 3
End Enum
Public Function ids_cal_resolution(FuseType As String, m_catename As String, ids_val As SiteDouble, ids_output As SiteDouble, resolution_output As Double, Optional showPrint = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim funcName As String:: funcName = "ids_cal_resolution" ''"auto_calc_IDS_Decimal"
    
    Dim m_str As String
    
    Dim ids_bool As New SiteBoolean
    
    Dim fuse_ids_resolution As Double
  
    Dim ids_max As New SiteDouble
    Dim IDS_resolution As New SiteDouble
    Dim ids_resolution_round As New SiteDouble
    Dim site As Variant 'Carter, 20240304
       
    ids_max = ids_val.Maximum(0)
    fuse_ids_resolution = auto_eFuse_GetIDSResolution(FuseType, m_catename)
    fuse_ids_resolution = fuse_ids_resolution * 0.001 ''''<NOTICE> update unit from mA to A
    IDS_resolution = ids_max.divide(fuse_ids_resolution)
    ids_resolution_round = IDS_resolution.Add(0.5).Truncate
    
    ids_bool = ids_resolution_round.Subtract(IDS_resolution).compare(GreaterThan, 0)
''-------Substitute the code in the belowing--------
    If ids_bool.Any(True) Then
        TheExec.sites.Selected = ids_bool
        ids_output = ids_resolution_round
        TheExec.sites.Selected = True
    End If
    
    If ids_bool.Any(False) Then
        TheExec.sites.Selected = ids_bool.LogicalNot
        ids_output = ids_resolution_round.Add(1)
        TheExec.sites.Selected = True
    End If

    ids_output = ids_output.Maximum(0)
       
    If (showPrint) Then
        For Each site In TheExec.sites.Active
            m_str = CStr(ids_val(site)) + "/" + CStr(fuse_ids_resolution)
            Debug.Print funcName + "......" + m_str + " = ids_resolution = " + CStr(IDS_resolution(site))
            Debug.Print funcName + "......" + m_str + " = ids_resolution_round = " + CStr(ids_resolution_round(site))
            Debug.Print funcName + "......" + m_str + " = " + CStr(ids_output(site))
            Debug.Print "---------------------------------------------------------------------" + vbCrLf
        Next site
    End If
    
    resolution_output = fuse_ids_resolution
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "ids_cal_resolution") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Find_Range_WaitTime(Range_List() As Double, Instrucment_Type As String, Optional MergedN As Long = 1, Optional PinIndex As Long = -1) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    
    Dim funcName As String:: funcName = "Find_Range_WaitTime"
    
    Dim var As Variant
    Dim SattleTime() As Double
    
    ReDim SattleTime(UBound(Range_List))
    
    If LCase(Instrucment_Type) = "hexvs" Then
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.01 Then
                SattleTime(var) = 100 * ms
            
            ElseIf Range_List(var) = 0.1 Then
                SattleTime(var) = 10 * ms + glb_HexVs_0p1range_waittime
                
            ElseIf Range_List(var) = 1 Then
                SattleTime(var) = 1 * ms
            
            ElseIf Range_List(var) >= 15 Then
                SattleTime(var) = 100 * us
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_WaitTime", "hexvs CurrangeRange not define!!!")
            End If
        Next var
        
    ElseIf LCase(Instrucment_Type) = "vhdvs" Then
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.000004 Then
                SattleTime(var) = 18 * ms
            
            ElseIf Range_List(var) = 0.00002 Then
                SattleTime(var) = 4 * ms
                
            ElseIf Range_List(var) = 0.0002 Then
                SattleTime(var) = 4 * ms
            
            ElseIf Range_List(var) = 0.002 Then
                SattleTime(var) = 3.5 * ms
                
            ElseIf Range_List(var) = 0.02 Then
                SattleTime(var) = 540 * us
                
            ElseIf Range_List(var) = 0.04 Then
                SattleTime(var) = 260 * us
                
            ElseIf Range_List(var) = 0.2 Then
                SattleTime(var) = 210 * us
            
            ElseIf Range_List(var) = 0.4 Then
                SattleTime(var) = 90 * us
            
            ElseIf Range_List(var) = 0.7 Then
                SattleTime(var) = 100 * us
                
            ElseIf Range_List(var) = 0.8 Then
                SattleTime(var) = 100 * us
                
            ElseIf Range_List(var) = 1.4 Then
                SattleTime(var) = 50 * us
                
            ElseIf Range_List(var) = 2.8 Then
                SattleTime(var) = 45 * us
                
            ElseIf Range_List(var) = 5.6 Then
                SattleTime(var) = 30 * us
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_WaitTime", "vhdvs CurrangeRange not define!!!")
            End If
        Next var
        
        
    ElseIf LCase(Instrucment_Type) = "dc-07" Then
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.000002 * MergedN Then
                SattleTime(var) = 6 * ms
            
            ElseIf Range_List(var) = 0.00002 * MergedN Then
                SattleTime(var) = 1.5 * ms
                
            ElseIf Range_List(var) = 0.0002 * MergedN Then
                SattleTime(var) = 1.4 * ms
            
            ElseIf Range_List(var) = 0.002 * MergedN Then
                SattleTime(var) = 11 * ms
                
            ElseIf Range_List(var) = 0.02 * MergedN Then
                SattleTime(var) = 1.5 * ms
            
            ElseIf Range_List(var) = 0.2 * MergedN Then
                SattleTime(var) = 260 * us
            
            ElseIf Range_List(var) >= 1 Then
                SattleTime(var) = 1.6 * ms
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_WaitTime", "dc-07 CurrangeRange not define!!!")
            End If
        Next var
        
    ElseIf LCase(Instrucment_Type) = "vsm" Then
        For var = 0 To UBound(Range_List)
            If Range_List(var) <= 10 Then
                SattleTime(var) = 6 * ms
                
            Else
                SattleTime(var) = 1.5 * ms
                
            End If
        Next var
     ElseIf LCase(Instrucment_Type) = "vs-5a" Or LCase(Instrucment_Type) = "vs-800ma" Then           '''210903 Merge UVS64 and UVS256HP to one case
    '''Hardware Specifications -> DC Instrument Specfication -> UltraVS64 -> Forcing and Metering Analog -> Current Metering DC
        
        If glb_TesterType = "UltraFLEXplus" And Flag_DC_ParsingRecoveryTable Then
                Call Calculate_SettleTimeByInstrument(PinIndex, Range_List, MergedN, SattleTime)
        Else
            
            If LCase(Instrucment_Type) = "vs-5a" Then
    
    '''Hardware Specifications -> DC Instrument Specfication -> UltraVS64 -> Forcing and Metering Analog -> Current Metering DC
                For var = 0 To UBound(Range_List)
                    If Range_List(var) = 0.00002 Then
                        SattleTime(var) = 40 * ms
                        
                    ElseIf Range_List(var) = 0.0002 Then
                        SattleTime(var) = 4 * ms
                        
                    ElseIf Range_List(var) = 0.002 Then
                        SattleTime(var) = 2 * ms
                        
                    ElseIf Range_List(var) = 0.02 Or Range_List(var) = 0.02 * MergedN Then
                        SattleTime(var) = 0.2 * ms
                            
                    ElseIf Range_List(var) = 0.2 Or Range_List(var) = 0.2 * MergedN Then
                        SattleTime(var) = 0.1 * ms
                            
                    ElseIf Range_List(var) = 0.5 Or Range_List(var) = 0.5 * MergedN Or Range_List(var) = 1 Or Range_List(var) = 1 * MergedN Or Range_List(var) = 5 Or Range_List(var) = 5 * MergedN Then
                        SattleTime(var) = 0.025 * ms
                    Else
                        Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_WaitTime", "vs-5a CurrangeRange not define!!!")
                    End If
                Next var
    
            ElseIf LCase(Instrucment_Type) = "vs-800ma" Then
            '''Hardware Specifications -> DC Instrument Specfication -> UltraVS256-HP -> Forcing and Metering Analog -> Current Metering DC
                For var = 0 To UBound(Range_List)
                    If Range_List(var) = 0.000004 Then
                        SattleTime(var) = 8.75 * ms
                        
                    ElseIf Range_List(var) = 0.00002 Or Range_List(var) = 0.002 Then
                        SattleTime(var) = 2 * ms
                        
                    ElseIf Range_List(var) = 0.0002 Or Range_List(var) = 0.02 Or Range_List(var) = 0.02 * MergedN Then
                        SattleTime(var) = 0.2 * ms
                        
                    ElseIf Range_List(var) = 0.2 Or Range_List(var) = 0.2 * MergedN Then
                        SattleTime(var) = 0.075 * ms
                         
                    ElseIf Range_List(var) = 0.8 Or Range_List(var) = 0.8 * MergedN Then
                        SattleTime(var) = 0.025 * ms
                    Else
                        Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_WaitTime", "vs-800ma CurrangeRange not define!!!")
                    End If
                Next var
            End If  'If LCase(Instrucment_Type) = "vs-5a" Then
        End If
    End If

    Find_Range_WaitTime = SattleTime

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Find_Range_WaitTime") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Find_Range_Accuracy(Range_List() As Double, Instrucment_Type As String, Optional MergedN As Long = 1) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    
    Dim funcName As String:: funcName = "Find_Range_WaitTime"
    
    Dim var As Variant
    Dim Accuracy() As Double
    
    ReDim Accuracy(UBound(Range_List))
    
    If LCase(Instrucment_Type) = "hexvs" Then
        
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.01 Then
                Accuracy(var) = Range_List(var) * 0.01 + 0.00005
                
            ElseIf Range_List(var) = 0.1 Then
                Accuracy(var) = Range_List(var) * 0.01 + 0.0005
                
            ElseIf Range_List(var) = 1 Then
                Accuracy(var) = Range_List(var) * 0.01 + 0.005
                
            ElseIf Range_List(var) >= 15 Then
                Accuracy(var) = Range_List(var) * 0.01 + (Range_List(var) \ 15) * 0.05
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_Accuracy", "hexvs CurrangeRange not define!!!")
            End If
        Next var
    ElseIf LCase(Instrucment_Type) = "vhdvs" Then
        
        For var = 0 To UBound(Range_List)
            
            If Range_List(var) = 0.000004 Then
                Accuracy(var) = Range_List(var) * 0.005 + 0.000000026
            ElseIf Range_List(var) = 0.00002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 0.00000012
            ElseIf Range_List(var) = 0.0002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 0.0000012
            ElseIf Range_List(var) = 0.002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 0.000012
            ElseIf Range_List(var) = 0.02 Then
                Accuracy(var) = Range_List(var) * 0.005 + 0.00012
            ElseIf Range_List(var) = 0.04 Then
                Accuracy(var) = Range_List(var) * 0.007 + 0.00024
            ElseIf Range_List(var) = 0.2 Then
                Accuracy(var) = Range_List(var) * 0.005 + 0.0012
            ElseIf Range_List(var) = 0.4 Then
                Accuracy(var) = Range_List(var) * 0.007 + 0.0024
            ElseIf Range_List(var) = 0.7 Then
                Accuracy(var) = Range_List(var) * 0.005 + 0.006
            ElseIf Range_List(var) = 0.8 Then
                Accuracy(var) = Range_List(var) * 0.005 + 0.006
            ElseIf Range_List(var) = 1.4 Then
                Accuracy(var) = Range_List(var) * 0.007 + 0.01
            ElseIf Range_List(var) = 2.8 Then
                Accuracy(var) = Range_List(var) * 0.007 + 0.02
            ElseIf Range_List(var) = 5.6 Then
                Accuracy(var) = Range_List(var) * 0.007 + 0.04
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_Accuracy", "vhdvs CurrangeRange not define!!!")
            End If
        Next var
    ElseIf LCase(Instrucment_Type) = "dc-07" Then
        
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.00002 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.002 + 0.000000075
            ElseIf Range_List(var) = 0.0002 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.002 + 0.0000004
            ElseIf Range_List(var) = 0.002 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.002 + 0.000004
            ElseIf Range_List(var) = 0.02 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.002 + 0.00004
            ElseIf Range_List(var) = 0.2 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.002 + 0.0004
            ElseIf Range_List(var) >= 1 Then
                Accuracy(var) = Range_List(var) * 0.002 + 0.008
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_Accuracy", "dc-07 CurrangeRange not define!!!")
            End If
        Next var
   ElseIf LCase(Instrucment_Type) = "vsm" Then
        
        For var = 0 To UBound(Range_List)
            If Range_List(var) <= 10 Then
                Accuracy(var) = Range_List(var) * 0.0025 + Range_List(var) * 0.007
            ElseIf (Range_List(var) Mod 11) = 0 Then
                Accuracy(var) = Range_List(var) * 0.0025 + (Range_List(var) \ 11) * 0.04
            ElseIf (Range_List(var) Mod 21) = 0 Then
                Accuracy(var) = Range_List(var) * 0.0025 + (Range_List(var) \ 21) * 0.08
            ElseIf (Range_List(var) Mod 51) = 0 Then
                Accuracy(var) = Range_List(var) * 0.0025 + (Range_List(var) \ 51) * 0.2
            ElseIf (Range_List(var) Mod 81) = 0 Then
                Accuracy(var) = Range_List(var) * 0.0025 + (Range_List(var) \ 81) * 0.32
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_Accuracy", "vsm CurrangeRange not define!!!")
            End If
        Next var
    ElseIf LCase(Instrucment_Type) = "vs-5a" Then
    '''Hardware Specifications -> DC Instrument Specfication -> UltraVS64 -> Forcing and Metering Analog -> Current Metering DC
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 5 Or Range_List(var) = 5 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.007 + 40 * MergedN * mA
            
            ElseIf Range_List(var) = 1 Or Range_List(var) = 1 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.007 + 8 * MergedN * mA
                
            ElseIf Range_List(var) = 0.5 Or Range_List(var) = 0.5 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.007 + 2.4 * MergedN * mA
                
            ElseIf Range_List(var) = 0.2 Or Range_List(var) = 0.2 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.005 + 1.2 * MergedN * mA
                
            ElseIf Range_List(var) = 0.02 Or Range_List(var) = 0.02 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.005 + 120 * MergedN * uA
                
            ElseIf Range_List(var) = 0.002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 12 * uA + 1 * (MergedN - 1) * uA
                
            ElseIf Range_List(var) = 0.0002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 1.2 * uA + 1 * (MergedN - 1) * uA
                
            ElseIf Range_List(var) = 0.00002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 120 * nA + 150 * (MergedN - 1) * nA
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_Accuracy", "vs-5a CurrangeRange not define!!!")
            End If
        Next var
    ElseIf LCase(Instrucment_Type) = "vs-800ma" Then
    '''Hardware Specifications -> DC Instrument Specfication -> UltraVS256-HP -> Forcing and Metering Analog -> Current Metering DC
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.8 Or Range_List(var) = 0.8 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.005 + 6 * MergedN * mA
                
            ElseIf Range_List(var) = 0.2 Or Range_List(var) = 0.2 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.005 + 1.2 * MergedN * mA
                
            ElseIf Range_List(var) = 0.02 Or Range_List(var) = 0.02 * MergedN Then
                Accuracy(var) = Range_List(var) * 0.005 + 120 * MergedN * uA
                
            ElseIf Range_List(var) = 0.002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 12 * uA
                
            ElseIf Range_List(var) = 0.0002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 1.2 * uA + 50 * (MergedN - 1) * nA
                
            ElseIf Range_List(var) = 0.00002 Then
                Accuracy(var) = Range_List(var) * 0.005 + 120 * nA + 50 * (MergedN - 1) * nA
                
            ElseIf Range_List(var) = 0.000004 Then
                Accuracy(var) = Range_List(var) * 0.005 + 36 * nA + 50 * (MergedN - 1) * nA
            Else
                Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_Accuracy", "vs-800ma CurrangeRange not define!!!")
            End If
        Next var
    
    End If
    
    Find_Range_Accuracy = Accuracy

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Find_Range_Accuracy") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20230425][All][Cylde] Add use in Plus instrument
Public Function InitIRange_Setup(ByRef pin_info As AutoRange_Info, _
                                ByRef WaitTime As Double, ByRef SattleTime As Double)
                                    
    Dim funcName As String:: funcName = "InitIRange_Steup"
    
    Dim j As Long
    
    Dim val As Double
    
    On Error GoTo errHandler
    
    Select Case UCase(pin_info.ChanMapType)
        Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VSM, glbConstIns_VS800MA, glbConstIns_VS5A:
            val = thehdw.DCVS.Pins(pin_info.PinName).CurrentLimit.Source.FoldLimit.level.value
            If val > thehdw.DCVS.Pins(pin_info.PinName).CurrentRange.max Then val = thehdw.DCVS.Pins(pin_info.PinName).CurrentRange.max
            For j = 0 To UBound(pin_info.Range_List) - 1
                If val >= pin_info.Range_List(j) And val <= pin_info.Range_List(j + 1) Then
                    pin_info.Init_step = j + 1
                    SattleTime = pin_info.WaitTime_List(j + 1) 'Michael fixed waittime index 20240527
                    thehdw.DCVS.Pins(pin_info.PinName).SetCurrentRanges val, val
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    Exit For
                End If
            Next j
            
        Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
            val = thehdw.DCVI.Pins(pin_info.PinName).Current
            If val > thehdw.DCVI.Pins(pin_info.PinName).CurrentRange.max Then val = thehdw.DCVI.Pins(pin_info.PinName).CurrentRange.max
            For j = 0 To UBound(pin_info.Range_List) - 1
                If val >= pin_info.Range_List(j) And val <= pin_info.Range_List(j + 1) Then
                    pin_info.Init_step = j + 1
                    SattleTime = pin_info.WaitTime_List(j + 1) 'Michael fixed waittime index 20240527
                    thehdw.DCVI.Pins(pin_info.PinName).SetCurrentAndRange val, val
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    Exit For
                End If
            Next j
    End Select
    
    If pin_info.Init_step = 0 Then
        pin_info.Init_step = UBound(pin_info.Range_List)
        SattleTime = pin_info.WaitTime_List(UBound(pin_info.Range_List))
        If SattleTime > WaitTime Then WaitTime = SattleTime
    End If
    
    Exit Function
    

errHandler:
    Call Print_Error_Message(Error_Info, LIB_DC, funcName)
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calculate_SettleTimeByInstrument(PinIndex As Long, Range_List() As Double, MergedN As Long, ByRef SettleTime() As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
Dim funName As String: funName = "Calculate_SettleTimeByInstrument"

    Dim dibcapIdx As Integer
    Dim c_CodeIdx As Integer
    Dim RangeIdx As Integer
    Dim SettleTimeIdx As Integer
    Dim InstrType As String
    Dim c_Code_final As Double
    Dim divideMerge As Double
    Dim TempCnt As Integer
    InstrType = PowerPin_range_ary(PinIndex).ChanMapType
    
    If Dict_Dibcap_PowerPinName.Exists(UCase(PowerPin_range_ary(PinIndex).PinName)) Then
    
        dibcapIdx = Dict_Dibcap_PowerPinName(UCase(PowerPin_range_ary(PinIndex).PinName))
        
        If LCase(PowerPin_range_ary(PinIndex).ChanMapType) = "vs-5a" Then  ''VS64
            c_Code_final = -1
            
            For c_CodeIdx = 0 To UBound(C_code_UVS64)
                If Dibcap(dibcapIdx).SingleEndCap >= C_code_UVS64(c_CodeIdx).min And Dibcap(dibcapIdx).SingleEndCap < C_code_UVS64(c_CodeIdx).max Then
                    c_Code_final = C_code_UVS64(c_CodeIdx).C_code * 16
                    
                    If c_Code_final > 335 Then c_Code_final = 335
                    
                    For SettleTimeIdx = 0 To UBound(DC_SettleTime_UVS64)
                        If c_Code_final >= DC_SettleTime_UVS64(SettleTimeIdx).BW_min And c_Code_final < DC_SettleTime_UVS64(SettleTimeIdx).BW_max Then
                            ''' update all current range settletime()
                            For RangeIdx = 0 To UBound(Range_List)
                            
                                If Dibcap(dibcapIdx).Merge <> 1 Then divideMerge = CDbl(Trim(Range_List(RangeIdx) / Dibcap(dibcapIdx).Merge))  'if merge<>1

                                If DC_SettleTime_UVS64(SettleTimeIdx).dict_CurrentRange.Exists(Range_List(RangeIdx)) Then
                                    If Dibcap(dibcapIdx).offset <> 0 Then
                                        SettleTime(RangeIdx) = DC_SettleTime_UVS64(SettleTimeIdx).dict_CurrentRange(Range_List(RangeIdx)) * ms + Dibcap(dibcapIdx).offset * ms
                                    Else
                                        SettleTime(RangeIdx) = DC_SettleTime_UVS64(SettleTimeIdx).dict_CurrentRange(Range_List(RangeIdx)) * ms  ''SettleTime(s)
                                    End If
                                Else
                                     '' if Range value isn't not exit in DC_SettleTime_UVS256HP(SettleTimeIdx).dict_CurrentRange that means it's Merge channel
                                    divideMerge = CDbl(Trim(Range_List(RangeIdx) / Dibcap(dibcapIdx).Merge))
                                    If Range_List(RangeIdx) = 40 Then divideMerge = 5
                                    
                                    ''' keep signelend range index in Range_list
                                    For TempCnt = 0 To UBound(Range_List)
                                        If Range_List(TempCnt) = divideMerge Then Exit For
                                    Next TempCnt
                                    
                                    If TempCnt = UBound(Range_List) + 1 Then ''if  divideMerge dosen't exist in Range_List then print error
                                        TheExec.Datalog.WriteComment "[Error] funName: " + funName + " Pin Name: " + PowerPin_range_ary(PinIndex).PinName + " current range is " + CStr(Range_List(RangeIdx)) + " dosen't exist in DC_SettleTime_UVS64"
                                        TheExec.ErrorLogMessage "[Error] funName: " + funName + " Pin Name: " + PowerPin_range_ary(PinIndex).PinName + " current range is " + CStr(Range_List(RangeIdx)) + " dosen't exist in DC_SettleTime_UVS64"
                                    End If
                                    
                                    If DC_SettleTime_UVS64(SettleTimeIdx).dict_CurrentRange.Exists(Range_List(TempCnt)) Then
                                        If Dibcap(dibcapIdx).offset <> 0 Then
                                            SettleTime(RangeIdx) = DC_SettleTime_UVS64(SettleTimeIdx).dict_CurrentRange(Range_List(RangeIdx)) * ms + Dibcap(dibcapIdx).offset * ms
                                        Else
                                            SettleTime(RangeIdx) = DC_SettleTime_UVS64(SettleTimeIdx).dict_CurrentRange(Range_List(TempCnt)) * ms
                                        End If
                                    Else
                                        TheExec.Datalog.WriteComment "[Error] funName: " + funName + " Pin Name: " + PowerPin_range_ary(PinIndex).PinName + " current range is " + CStr(Range_List(RangeIdx)) + " dosen't exist in DC_SettleTime_UVS64"
                                    End If
                                End If

                            Next RangeIdx
                            Exit For
                        End If
                    Next SettleTimeIdx
                    
                    Exit For
                End If
            Next c_CodeIdx
                    
        ElseIf LCase(PowerPin_range_ary(PinIndex).ChanMapType) = "vs-800ma" Then   ''UVS256-HP
            c_Code_final = -1
            
            For c_CodeIdx = 0 To UBound(C_code_UVS256HP)
                If Dibcap(dibcapIdx).SingleEndCap >= C_code_UVS256HP(c_CodeIdx).min And Dibcap(dibcapIdx).SingleEndCap < C_code_UVS256HP(c_CodeIdx).max Then
                    c_Code_final = C_code_UVS256HP(c_CodeIdx).C_code * 16
                    
                    If c_Code_final > 255 Then c_Code_final = 240
                    
                    For SettleTimeIdx = 0 To UBound(DC_SettleTime_UVS256HP)
                        If c_Code_final >= DC_SettleTime_UVS256HP(SettleTimeIdx).BW_min And c_Code_final < DC_SettleTime_UVS256HP(SettleTimeIdx).BW_max Then
                            ''' update all current range settletime()
                            For RangeIdx = 0 To UBound(Range_List)
                                
                                If Dibcap(dibcapIdx).Merge <> 1 Then divideMerge = CDbl(Trim(Range_List(RangeIdx) / Dibcap(dibcapIdx).Merge))

                                If DC_SettleTime_UVS256HP(SettleTimeIdx).dict_CurrentRange.Exists(Range_List(RangeIdx)) Then
                                    If Dibcap(dibcapIdx).offset <> 0 Then
                                        SettleTime(RangeIdx) = DC_SettleTime_UVS64(SettleTimeIdx).dict_CurrentRange(Range_List(RangeIdx)) * ms + Dibcap(dibcapIdx).offset * ms
                                   Else
                                        SettleTime(RangeIdx) = DC_SettleTime_UVS256HP(SettleTimeIdx).dict_CurrentRange(Range_List(RangeIdx)) * ms
                                    End If
                                Else
                                '' if Range value isn't not exit in DC_SettleTime_UVS256HP(SettleTimeIdx).dict_CurrentRange that means it's Merge channel
                                    divideMerge = CDbl(Trim(Range_List(RangeIdx) / Dibcap(dibcapIdx).Merge))
                                    ''' keep signelend range index in Range_list
                                    For TempCnt = 0 To UBound(Range_List)
                                        If Range_List(TempCnt) = divideMerge Then Exit For
                                    Next TempCnt
                                    
                                    If TempCnt = UBound(Range_List) + 1 Then ''if  divideMergedosen't exist in Range_List then print error
                                        TheExec.Datalog.WriteComment "[Error] funName: " + funName + " Pin Name: " + PowerPin_range_ary(PinIndex).PinName + " current range is " + CStr(Range_List(RangeIdx)) + " dosen't exist in DC_SettleTime_UVS64"
                                        TheExec.ErrorLogMessage "[Error] funName: " + funName + " Pin Name: " + PowerPin_range_ary(PinIndex).PinName + " current range is " + CStr(Range_List(RangeIdx)) + " dosen't exist in DC_SettleTime_UVS64"
                                    End If

                                    If DC_SettleTime_UVS256HP(SettleTimeIdx).dict_CurrentRange.Exists(Range_List(TempCnt)) Then
                                        If Dibcap(dibcapIdx).offset <> 0 Then
                                            SettleTime(RangeIdx) = DC_SettleTime_UVS64(SettleTimeIdx).dict_CurrentRange(Range_List(RangeIdx)) * ms + Dibcap(dibcapIdx).offset * ms
                                        Else
                                            SettleTime(RangeIdx) = DC_SettleTime_UVS256HP(SettleTimeIdx).dict_CurrentRange(Range_List(TempCnt)) * ms
                                        End If
                                    Else
                                        TheExec.Datalog.WriteComment "[Error] funName: " + funName + " Pin Name: " + PowerPin_range_ary(PinIndex).PinName + " current range is " + CStr(Range_List(RangeIdx)) + " dosen't exist in DC_SettleTime_UVS256HP"
                                    End If
                                End If
                                
                            Next RangeIdx
                            Exit For
                        End If
                    Next SettleTimeIdx
                    
                    Exit For
                
                End If
            Next c_CodeIdx
            
        End If
    
''''    Else        '''Not Dict_Dibcap_PowerPinName.Exists(UCase(PowerPin_range_ary(pinIndex).pinName))
''''        ''' UVS64
''''        If LCase(PowerPin_range_ary(pinIndex).ChanMapType) = "vs-5a" Then
''''        '''Hardware Specifications -> DC Instrument Specfication -> UltraVS64 -> Forcing and Metering Analog -> Current Metering DC
''''            For RangeIdx = 0 To UBound(Range_List)
''''                If Range_List(RangeIdx) = 0.00002 Then
''''                    SettleTime(RangeIdx) = 40 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.0002 Then
''''                    SettleTime(RangeIdx) = 4 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.002 Then
''''                    SettleTime(RangeIdx) = 2 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.02 Or Range_List(RangeIdx) = 0.02 * MergedN Then
''''                    SettleTime(RangeIdx) = 0.2 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.2 Or Range_List(RangeIdx) = 0.2 * MergedN Then
''''                    SettleTime(RangeIdx) = 0.1 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.5 Or Range_List(RangeIdx) = 0.5 * MergedN Or Range_List(RangeIdx) = 1 Or Range_List(RangeIdx) = 1 * MergedN Or Range_List(RangeIdx) = 5 Or Range_List(RangeIdx) = 5 * MergedN Then
''''                    SettleTime(RangeIdx) = 0.025 * ms
''''                End If
''''            Next RangeIdx
''''        '''UVS256-HP
''''        ElseIf LCase(PowerPin_range_ary(pinIndex).ChanMapType) = "vs-800ma" Then
''''        '''Hardware Specifications -> DC Instrument Specfication -> UltraVS256-HP -> Forcing and Metering Analog -> Current Metering DC
''''            For RangeIdx = 0 To UBound(Range_List)
''''                If Range_List(RangeIdx) = 0.000004 Then
''''                    SettleTime(RangeIdx) = 8.75 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.00002 Or Range_List(RangeIdx) = 0.002 Then
''''                    SettleTime(RangeIdx) = 2 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.0002 Or Range_List(RangeIdx) = 0.02 Or Range_List(RangeIdx) = 0.02 * MergedN Then
''''                    SettleTime(RangeIdx) = 0.2 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.2 Or Range_List(RangeIdx) = 0.2 * MergedN Then
''''                    SettleTime(RangeIdx) = 0.075 * ms
''''
''''                ElseIf Range_List(RangeIdx) = 0.8 Or Range_List(RangeIdx) = 0.8 * MergedN Then
''''                    SettleTime(RangeIdx) = 0.025 * ms
''''                End If
''''            Next RangeIdx
''''
''''        End If
        
    End If

Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Calculate_SettleTimeByInstrument") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'20210831 add for  C code structure
Public Function Init_UVS256HP_C_code()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funName As String: funName = "Init_UVS256HP_C_code"
 '''********define current range********
''' C code          Unmerged Min        Max
'''     0                       0                               0.05
'''     1                       0.05                        0.083
'''     2                       0.083                       0.138
'''     3                       0.138                       0.229
'''     4                       0.229                       0.38
'''     5                       0.38                        0.63
'''     6                       0.63                        1.1
'''     7                       1.1                         1.7
'''     8                       1.7                         2.9
'''     9                       2.9                         4.8
'''     10                  4.8                             7.9
'''     11                  7.9                             13
'''     12                  13                              22
'''     13                  22                              36
'''     14                  36                              60
'''     15                  60                              100
'''
'-------------UVS256-HP-----------------
    C_code_UVS256HP(0).min = 0
    C_code_UVS256HP(0).max = 0.05
    C_code_UVS256HP(0).C_code = 0
    
    C_code_UVS256HP(1).min = 0.05
    C_code_UVS256HP(1).max = 0.083
    C_code_UVS256HP(1).C_code = 1
    
    C_code_UVS256HP(2).min = 0.083
    C_code_UVS256HP(2).max = 0.138
    C_code_UVS256HP(2).C_code = 2
    
    C_code_UVS256HP(3).min = 0.138
    C_code_UVS256HP(3).max = 0.229
    C_code_UVS256HP(3).C_code = 3
    
    C_code_UVS256HP(4).min = 0.229
    C_code_UVS256HP(4).max = 0.38
    C_code_UVS256HP(4).C_code = 4
    
    C_code_UVS256HP(5).min = 0.38
    C_code_UVS256HP(5).max = 0.63
    C_code_UVS256HP(5).C_code = 5
    
    C_code_UVS256HP(6).min = 0.63
    C_code_UVS256HP(6).max = 1.1
    C_code_UVS256HP(6).C_code = 6
    
    C_code_UVS256HP(7).min = 1.1
    C_code_UVS256HP(7).max = 1.7
    C_code_UVS256HP(7).C_code = 7
    
    C_code_UVS256HP(8).min = 1.7
    C_code_UVS256HP(8).max = 2.9
    C_code_UVS256HP(8).C_code = 8
    
    C_code_UVS256HP(9).min = 2.9
    C_code_UVS256HP(9).max = 4.8
    C_code_UVS256HP(9).C_code = 9
    
    C_code_UVS256HP(10).min = 4.8
    C_code_UVS256HP(10).max = 7.9
    C_code_UVS256HP(10).C_code = 10
    
    C_code_UVS256HP(11).min = 7.9
    C_code_UVS256HP(11).max = 13
    C_code_UVS256HP(11).C_code = 11
    
    C_code_UVS256HP(12).min = 13
    C_code_UVS256HP(12).max = 22
    C_code_UVS256HP(12).C_code = 12
    
    C_code_UVS256HP(13).min = 22
    C_code_UVS256HP(13).max = 36
    C_code_UVS256HP(13).C_code = 13
    
    C_code_UVS256HP(14).min = 36
     C_code_UVS256HP(14).max = 60
     C_code_UVS256HP(14).C_code = 14
    
    C_code_UVS256HP(15).min = 60
    C_code_UVS256HP(15).max = 100
    C_code_UVS256HP(15).C_code = 15
   

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Init_UVS256HP_C_code") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'20210810 add for parsing UVS256 Table
Public Function Init_UVS256HP_SettleTime()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funName As String: funName = "Init_UVS256HP_SettleTime"
'''================================================================================================================================================='''
'''parsing UVS256-HP settle time table
''BW setting      UVS256HP Current Range, settle time in uS
'''min max     4uA           20uA        200uA       2mA     20mA        200mA       800mA
'''0     15          1,880               376         176             376         176            75              25           --- [0]
'''16  31           8,800            1,760       176             376         176             75              25           --- [1]
'''32  47           8,800            1,760       176             376         176             75              25           --- [1]
'''48  63            40,000       8,000          800             376         176             75              25           --- [2]
'''64  79            40,000      8,000           800             376         176             75              25           --- [2]
'''80  95            40,000      8,000           800             376         176             75              25           --- [2]
'''96  111       188,000     37,600      3,760              376          176            75               25           --- [3]
'''112 127     188,000     37,600      3,760             376         176             75              25           --- [3]
'''128 143     188,000     37,600      3,760             376         176             75              25           --- [3]
'''144 159     880,000     176,000     17,600      1,760       176             75              25           --- [4]
'''160 175     880,000     176,000     17,600      1,760       176             75              25           --- [4]
'''176 191     880,000     176,000     17,600      1,760       176             75              25           --- [4]
'''192 207     880,000     176,000     17,600      1,760       176             75              25           --- [4]
'''208 223     880,000     176,000     17,600      1,760       176             75              25           --- [4]
'''224 239     880,000     176,000     17,600      1,760       176             75              25           --- [4]
'''240 255     880,000     176,000     17,600      1,760       176             75              25           --- [4]
'''================================================================================================================================================='''
    '''*****************************************************************'''
    DC_SettleTime_UVS256HP(0).BW_min = 0
    DC_SettleTime_UVS256HP(0).BW_max = 15
    DC_SettleTime_UVS256HP(0).dict_CurrentRange.Add 0.000004, 1.88 '// 0.000004A, 1.88ms
    DC_SettleTime_UVS256HP(0).dict_CurrentRange.Add 0.00002, 0.376 '//0.00002A, 0.376 ms
    DC_SettleTime_UVS256HP(0).dict_CurrentRange.Add 0.0002, 0.176
    DC_SettleTime_UVS256HP(0).dict_CurrentRange.Add 0.002, 0.376
    DC_SettleTime_UVS256HP(0).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS256HP(0).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS256HP(0).dict_CurrentRange.Add 0.8, 0.025
    
    
    DC_SettleTime_UVS256HP(1).BW_min = 16
    DC_SettleTime_UVS256HP(1).BW_max = 47
    DC_SettleTime_UVS256HP(1).dict_CurrentRange.Add 0.000004, 8.8
    DC_SettleTime_UVS256HP(1).dict_CurrentRange.Add 0.00002, 1.76
    DC_SettleTime_UVS256HP(1).dict_CurrentRange.Add 0.0002, 0.176
    DC_SettleTime_UVS256HP(1).dict_CurrentRange.Add 0.002, 0.376
    DC_SettleTime_UVS256HP(1).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS256HP(1).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS256HP(1).dict_CurrentRange.Add 0.8, 0.025
    
    DC_SettleTime_UVS256HP(2).BW_min = 48
    DC_SettleTime_UVS256HP(2).BW_max = 95
    DC_SettleTime_UVS256HP(2).dict_CurrentRange.Add 0.000004, 40
    DC_SettleTime_UVS256HP(2).dict_CurrentRange.Add 0.00002, 8
    DC_SettleTime_UVS256HP(2).dict_CurrentRange.Add 0.0002, 0.8
    DC_SettleTime_UVS256HP(2).dict_CurrentRange.Add 0.002, 0.376
    DC_SettleTime_UVS256HP(2).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS256HP(2).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS256HP(2).dict_CurrentRange.Add 0.8, 0.025
    
    DC_SettleTime_UVS256HP(3).BW_min = 96
    DC_SettleTime_UVS256HP(3).BW_max = 143
    DC_SettleTime_UVS256HP(3).dict_CurrentRange.Add 0.000004, 188
    DC_SettleTime_UVS256HP(3).dict_CurrentRange.Add 0.00002, 37.6
    DC_SettleTime_UVS256HP(3).dict_CurrentRange.Add 0.0002, 3.76
    DC_SettleTime_UVS256HP(3).dict_CurrentRange.Add 0.002, 0.376
    DC_SettleTime_UVS256HP(3).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS256HP(3).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS256HP(3).dict_CurrentRange.Add 0.8, 0.025
    
    DC_SettleTime_UVS256HP(4).BW_min = 144
    DC_SettleTime_UVS256HP(4).BW_max = 255
    DC_SettleTime_UVS256HP(4).dict_CurrentRange.Add 0.000004, 880
    DC_SettleTime_UVS256HP(4).dict_CurrentRange.Add 0.00002, 176
    DC_SettleTime_UVS256HP(4).dict_CurrentRange.Add 0.0002, 17.6
    DC_SettleTime_UVS256HP(4).dict_CurrentRange.Add 0.002, 1.76
    DC_SettleTime_UVS256HP(4).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS256HP(4).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS256HP(4).dict_CurrentRange.Add 0.8, 0.025

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Init_UVS256HP_SettleTime") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


'20210831 add for  C code structure
Public Function Init_UVS64_C_code()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funName As String: funName = "Init_UVS64_C_code"
 '''********define current range********
''' C code          Unmerged Min        Max
'''     0                       0                               0.05
'''     1                   0.05                            0.083
'''     2                   0.083                           0.138
'''     3                   0.138                           0.229
'''     4                   0.229                           0.38
'''     5                   0.38                            0.63
'''     6                   0.63                            1.1
'''     7                   1.1                             1.7
'''     8                   1.7                             2.9
'''     9                   2.9                             4.8
'''     10              4.8                             7.9
'''     11              7.9                             13
'''     12              13                              22
'''     13              22                              36
'''     14              36                              60
'''     15              60                              100
'''     16              100                             160
'''     17              160                             250
'''     18              250                             400
'''     19              400                             640
'''     20              640                             1000

'-------------VS64-----------------

    C_code_UVS64(0).min = 0
    C_code_UVS64(0).max = 0.05
    C_code_UVS64(0).C_code = 0
    
    C_code_UVS64(1).min = 0.05
    C_code_UVS64(1).max = 0.083
    C_code_UVS64(1).C_code = 1
    
    C_code_UVS64(2).min = 0.083
    C_code_UVS64(2).max = 0.138
    C_code_UVS64(2).C_code = 2
    
    C_code_UVS64(3).min = 0.138
    C_code_UVS64(3).max = 0.229
    C_code_UVS64(3).C_code = 3
    
    C_code_UVS64(4).min = 0.229
    C_code_UVS64(4).max = 0.38
    C_code_UVS64(4).C_code = 4
    
    C_code_UVS64(5).min = 0.38
    C_code_UVS64(5).max = 0.63
    C_code_UVS64(5).C_code = 5
    
    C_code_UVS64(6).min = 0.63
    C_code_UVS64(6).max = 1.1
    C_code_UVS64(6).C_code = 6
    
    C_code_UVS64(7).min = 1.1
    C_code_UVS64(7).max = 1.7
    C_code_UVS64(7).C_code = 7
    
    C_code_UVS64(8).min = 1.7
    C_code_UVS64(8).max = 2.9
    C_code_UVS64(8).C_code = 8
    
    C_code_UVS64(9).min = 2.9
    C_code_UVS64(9).max = 4.8
    C_code_UVS64(9).C_code = 9
    
    C_code_UVS64(10).min = 4.8
    C_code_UVS64(10).max = 7.9
    C_code_UVS64(10).C_code = 10
    
    C_code_UVS64(11).min = 7.9
    C_code_UVS64(11).max = 13
    C_code_UVS64(11).C_code = 11
    
    C_code_UVS64(12).min = 13
    C_code_UVS64(12).max = 22
    C_code_UVS64(12).C_code = 12
    
    C_code_UVS64(13).min = 22
    C_code_UVS64(13).max = 36
    C_code_UVS64(13).C_code = 13
    
    C_code_UVS64(14).min = 36
    C_code_UVS64(14).max = 60
    C_code_UVS64(14).C_code = 14
    
    C_code_UVS64(15).min = 60
    C_code_UVS64(15).max = 100
    C_code_UVS64(15).C_code = 15
    
    C_code_UVS64(16).min = 100
    C_code_UVS64(16).max = 160
    C_code_UVS64(16).C_code = 16
    
    C_code_UVS64(17).min = 160
    C_code_UVS64(17).max = 250
    C_code_UVS64(17).C_code = 17
    
    C_code_UVS64(18).min = 250
    C_code_UVS64(18).max = 400
    C_code_UVS64(18).C_code = 18
    
    C_code_UVS64(19).min = 400
     C_code_UVS64(19).max = 640
     C_code_UVS64(19).C_code = 19
    
    C_code_UVS64(20).min = 640
    C_code_UVS64(20).max = 1000
    C_code_UVS64(20).C_code = 20

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Init_UVS64_C_code") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function Init_UVS64_SettleTime()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funName As String: funName = "Init_UVS64_SettleTime"

'''================================================================================================================================================='''
'''parsing UVS64 settle time table
''BW setting      UVS256HP Current Range, settle time in uS
'''min  max     4uA             20uA        200uA       2mA         20mA            200mA       500mA           1A          5A
'''0        15          1,880           376             176             376             176                 75              25                  25          25             --- [0]
'''16       31          8,800       1,760              176             376             176                 75              25               25          25           --- [1]
'''32       47          8,800       1,760             176             376             176                 75              25                25          25           --- [1]
'''48       63      40,000          8,000           800             376             176                 75              25                  25          25           --- [2]
'''64       79      40,000          8,000           800             376             176                 75              25                  25          25           --- [2]
'''80       95      40,000          8,000           800             376             176                 75              25                  25          25           --- [2]
'''96      111      188,000     37,600      3,760            376             176                 75              25                  25          25           --- [3]
'''112     127     188,000     37,600      3,760           376             176                 75              25                  25          25           --- [3]
'''128      143     188,000     37,600      3,760           376             176                 75              25                  25          25           --- [3]
'''144      159     880,000     176,000     17,600      1,760           176                 75              25                  25          25           --- [4]
'''160      175     880,000     176,000     17,600      1,760           176                75              25                  25          25          --- [4]
'''176      191     880,000     176,000     17,600      1,760           176                75              25                  25          25          --- [4]
'''192      207      880,000     176,000     17,600      1,760           176                75              25                  25          25          --- [4]
'''208      223      880,000     176,000     17,600      1,760           176                75              25                  25          25          --- [4]
'''224      239     880,000     176,000     17,600      1,760           176                75              25                  25          25          --- [4]
'''240      255      880,000     176,000     17,600      1,760           176                75              25                 25          25          --- [4]
'''256      258                                                                                                                                                      25            25          --- [4]
'''272      274                                                                                                                                                       25            25          --- [4]
'''288      290                                                                                                                                                       25            25          --- [4]
'''304      306                                                                                                                                                       25            25          --- [4]
'''320      322                                                                                                                                                       25            25          --- [4]
'''335      335     880,000     176,000     17,600        1,760         176                 75              25                  25          25          --- [4]

'''================================================================================================================================================='''
'''*****************************************************************'''
    DC_SettleTime_UVS64(0).BW_min = 0
    DC_SettleTime_UVS64(0).BW_max = 15
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 0.000004, 1.88 ''//0.000004A, 1.88ms
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 0.00002, 0.376 '//0.00002A, 0.376 ms
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 0.0002, 0.176
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 0.002, 0.376
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 0.5, 0.025
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 1, 0.025
    DC_SettleTime_UVS64(0).dict_CurrentRange.Add 5, 0.025
    
    
    DC_SettleTime_UVS64(1).BW_min = 16
    DC_SettleTime_UVS64(1).BW_max = 47
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 4, 8.8
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 0.00002, 1.76
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 0.0002, 0.176
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 0.002, 0.376
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 0.5, 0.025
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 1, 0.025
    DC_SettleTime_UVS64(1).dict_CurrentRange.Add 5, 0.025
       
    
    DC_SettleTime_UVS64(2).BW_min = 48
    DC_SettleTime_UVS64(2).BW_max = 95
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 4, 40
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 0.00002, 8
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 0.0002, 0.8
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 0.002, 0.0376
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 0.5, 0.025
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 1, 0.025
    DC_SettleTime_UVS64(2).dict_CurrentRange.Add 5, 0.025
    
    DC_SettleTime_UVS64(3).BW_min = 96
    DC_SettleTime_UVS64(3).BW_max = 143
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 4, 188
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 0.00002, 37.6
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 0.0002, 3.76
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 0.002, 0.376
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 0.5, 0.025
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 1, 0.025
    DC_SettleTime_UVS64(3).dict_CurrentRange.Add 5, 0.025
    
    DC_SettleTime_UVS64(4).BW_min = 144
    DC_SettleTime_UVS64(4).BW_max = 335
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 4, 880
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 0.00002, 176
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 0.0002, 17.6
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 0.002, 1.76
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 0.02, 0.176
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 0.2, 0.075
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 0.5, 0.025
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 1, 0.025
    DC_SettleTime_UVS64(4).dict_CurrentRange.Add 5, 0.025

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Init_UVS64_SettleTime") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

''210903 For initial settle time mapping table and c_code mapping table
Public Function Initial_UVS256HP_UVS64_settleTime_CCode()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funName As String: funName = "Initial_UVS256HP_UVS64_settleTime_CCode"

    Init_UVS256HP_SettleTime
    Init_UVS64_SettleTime
    Init_UVS256HP_C_code
    Init_UVS64_C_code

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Initial_UVS256HP_UVS64_settleTime_CCode") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Parsing_dibcap_table()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funName As String: funName = "Parsing_dibcap_table"

    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim col_load As Integer
    Dim i As Long
    Dim Row As Long, Col As Long
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim isSheetFound As Boolean
    Dim powerpin_col As Integer
    Dim Merge_col As Integer
    Dim CapValue_col As Integer
    Dim Search_headline As Boolean
    Dim parsing_table As Boolean
    Dim offset_col As Integer
'''================================================================================================================================================='''
'''Sheets "dibcap_cp"
'''================================================================================================================================================='''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    If UCase(TheExec.CurrentJob) Like UCase("*CP*") Or UCase(TheExec.CurrentJob) Like UCase("*WLFT*") Then '''//job name
        sheetName = "dibcap_cp"
    ElseIf UCase(TheExec.CurrentJob) Like UCase("*FT*") Then '''//job name
        sheetName = "dibcap_ft"
    Else
        TheExec.Datalog.WriteComment "Error in Caculate_RecoveryTimes, can't define the job name"
        TheExec.ErrorLogMessage "Error in Caculate_RecoveryTimes, can't define the job name"
    End If
    
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    '''*****************************************************************'''
    If isSheetFound = True Then
        Application.ScreenUpdating = False
        '''//init
        i = 0
        powerpin_col = -1
        Merge_col = -1
        offset_col = -1
        CapValue_col = -1
        Search_headline = True
        parsing_table = True
       
        For Row = 1 To MaxRow
            ReDim Preserve Dibcap(i)
            If Search_headline = True Then
                For Col = 1 To maxcol
                    If Trim(LCase(ws_def.Cells(Row, Col).value)) Like LCase("power*pin*") Then '''//powerpin
                        powerpin_col = Col
                    ElseIf Trim(LCase(ws_def.Cells(Row, Col).value)) Like LCase("*Merge*") Then '''//Merge
                        Merge_col = Col
                    ElseIf Trim(LCase(ws_def.Cells(Row, Col).value)) Like LCase("*CapValue*") Then '''//CapValue
                        CapValue_col = Col
                    ElseIf Trim(LCase(ws_def.Cells(Row, Col).value)) Like LCase("*offset*") Then '''//offset
                        offset_col = Col
                    End If
                Next Col
       
                If powerpin_col <> -1 And Merge_col <> -1 And CapValue_col <> -1 And offset_col <> -1 Then
                    Search_headline = False
                End If
            End If
            
            If UCase(ws_def.Cells(Row, powerpin_col).value) Like UCase("*VDD*") Then
                Dibcap(i).powerPin = UCase(Trim(ws_def.Cells(Row, powerpin_col).value))
                Dibcap(i).Merge = CInt(Trim(ws_def.Cells(Row, Merge_col).value))
                Dibcap(i).CapValue = CDbl(Trim(ws_def.Cells(Row, CapValue_col).value))
                If Dibcap(i).offset <> 0 Then
                    Dibcap(i).offset = CDbl(Trim(ws_def.Cells(Row, offset_col).value))
                End If
                Dibcap(i).SingleEndCap = Dibcap(i).CapValue / Dibcap(i).Merge  ''SingleEndCap=CapValue for each pin
                Dict_Dibcap_PowerPinName.Add Dibcap(i).powerPin, i           ''arrange the power pin in number
                i = i + 1
            End If
            If UCase(ws_def.Cells(Row, powerpin_col).value) Like UCase("*End*") Then Exit For
        Next Row
        Application.ScreenUpdating = True
        
        Flag_DC_ParsingRecoveryTable = True
        
    End If '''If isSheetFound = True
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Parsing_dibcap_table") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Disable_FRC_Pins(FRC_Pins As String, Realy_Pins As String, DisconnectClock As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'change XI0/XO0 clock to MCG for IDS pattern run
 
    thehdw.Utility.Pins(Realy_Pins).State = tlUtilBitOn
    If DisconnectClock Then thehdw.Digital.Pins(FRC_Pins).Disconnect
    
    With thehdw.PPMU(FRC_Pins)
        .ForceV 0
        .Connect
        .Gate = tlOn
    End With

    Wait 0.005
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Disable_FRC_Pins") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Enable_FRC_Pins(FRC_Pins As String, Relay_Pins As String, DisconnectClock As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    '****CW 20210308 change XI0/XO0 clock to MCG for IDS pattern run
        
    With thehdw.PPMU(FRC_Pins)
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Utility.Pins(Relay_Pins).State = tlUtilBitOff
    If DisconnectClock Then thehdw.Digital.Pins(FRC_Pins).Connect
    Wait 0.005

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Enable_FRC_Pins") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'Public Sub GetSheetInfo(sheetName As String, MaxRow As Long, maxcol As Long, ParsingArr() As Variant)
'
'    Dim MySheet As New Worksheet
'    Dim m_StartRow As Long:: m_StartRow = 1
'
'    Set MySheet = Sheets(sheetName)            ''Active sheet
'    MySheet.Activate
'
'    MaxRow = MySheet.UsedRange.Rows.Count    ''Get the max used rows and columns of BDF.
'    maxcol = MySheet.UsedRange.Columns.Count
'
'    ParsingArr = MySheet.range(Cells(m_StartRow, 1), Cells(MaxRow, maxcol)).value
'
'End Sub

''20230223: Modidfied to fix VBT error.
' [20231003][All][Tank] modify after Chihome review
Public Function ParsePreCondition()
    Dim sFuncName As String:: sFuncName = "ParsePreCondition"
    Dim sheetName As String
    Dim SheetMaxRow As Long
    Dim SheetMaxCol As Long
    Dim ParsingArr() As Variant
    Dim MySheet As New Worksheet
    Dim m_StartRow As Long:: m_StartRow = 1
    Dim m_ParsingArr() As Variant
    Dim m_PatLoop As Long: m_PatLoop = 2
    Dim m_XpinIndex As Long
    Dim m_UnusedPinIndex As Long
    Dim m_AllPinIndex As Long

    Dim patindex As Long: patindex = 0
    Dim XpinIndex As Long
    Dim m_loop As Long

    Dim EndIndex As Long
    Dim m_start As Long
    Dim m_end As Long

    Dim m_PinStr As Variant
    Dim m_AllPinStr As Variant
    Dim dic_nWireLevelName As New Dictionary
    Dim LevelLoop As Long
    Dim CellTmp As String
    Dim tmp As New Dictionary
    Dim PinIndex As Long: PinIndex = 0
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim Total_Pin_cnt As Long
    Dim PinLoop As Long
    Dim PinValue As Long
    Dim AllSheetNames() As String
    Dim i As Long
    Dim Levels_nWire_exist As Boolean
    Dim Levels_IDS_exist As Boolean
    Dim s_ErrorStatement As String
    Dim dic_PreComditionPatName As New Dictionary
    Dim asTempPinName() As String
    Dim vPin As Variant
    Dim sPinName As String
    Dim sTempPinList As String

    Const cst_PatName = 1
    Const cst_PatVersion = 2
    Const cst_XPin = 3
    Const cst_UnusedPin = 4
    
    On Error GoTo errHandler

    If gl_isParsePreConditionDone = False Then

        gl_isCheckPreConditionVerDone = False     'inital flag
        gl_ePreConditionError = PreConditionError.SheetNotExist  'inital flag

        ''-------------------------
        ''Parse Levels_nWire Sheet
        ''-------------------------
        sheetName = "Levels_nWire"
        
        AllSheetNames() = TheExec.job.GetAllSheetNames
        
        For i = 0 To UBound(AllSheetNames)
            If LCase(AllSheetNames(i)) Like LCase("*Levels_nWire*") Then
                Call GetSheetInfo(sheetName, SheetMaxRow, SheetMaxCol, ParsingArr)
        
                Const PinGroupIndex = 2
                m_loop = 4
                Total_Pin_cnt = 0
                Do
                    If tmp.Exists(ParsingArr(m_loop, PinGroupIndex)) Or ParsingArr(m_loop, PinGroupIndex) = "" Then
                    Else
                        tmp.Add ParsingArr(m_loop, PinGroupIndex), ParsingArr(m_loop, PinGroupIndex)
                        TheExec.DataManager.DecomposePinList ParsingArr(m_loop, PinGroupIndex), Pin_Ary, Pin_Cnt
    
                        For PinLoop = 0 To Pin_Cnt - 1
                            If dic_nWireLevelName.Exists(Pin_Ary(PinLoop)) Then
                                dic_nWireLevelName.Remove Pin_Ary(PinLoop)
                            Else
                            End If
                            dic_nWireLevelName.Add Pin_Ary(PinLoop), (PinLoop + Total_Pin_cnt)
                            
                        Next PinLoop
                        Total_Pin_cnt = Pin_Cnt + Total_Pin_cnt
                        PinIndex = PinIndex + 1
                    End If
                    m_loop = m_loop + 1
                Loop Until (m_loop = SheetMaxRow + 1)
                Levels_nWire_exist = True
            
            End If
        Next i
        If Levels_nWire_exist = False Then
            s_ErrorStatement = sheetName & " doesn't exist in current program. Error!!!"
            Call Print_Error_Message(Warning_Info, LIB_DC, sFuncName, s_ErrorStatement)      '20221205 Tank try use print error
        End If
        ''-------------------------
        ''Parse PreCondition Table
        ''-------------------------
        sheetName = "PreCondition"
        
        
        If Find_Sheet(sheetName) = True Then
       
            Call GetSheetInfo(sheetName, SheetMaxRow, SheetMaxCol, ParsingArr)
            
    
            ''----------------
            ''Get Pattern name
            ''----------------
            Do
                If ParsingArr(m_PatLoop, cst_PatName) <> "" Then
                    If dic_PreComditionPatName.Exists(ParsingArr(m_PatLoop, cst_PatName)) Then
                        gl_ePreConditionError = PreConditionError.PatRepetition
                        Exit Function
                    Else
                        dic_PreComditionPatName.Add ParsingArr(m_PatLoop, cst_PatName), patindex
                    End If
                    
                    ReDim Preserve PreConditionInfo(patindex)
                    PreConditionInfo(patindex).PatternName = ParsingArr(m_PatLoop, cst_PatName)
                    PreConditionInfo(patindex).PatCell = m_PatLoop
                    If ParsingArr(m_PatLoop, cst_PatVersion) <> "" Then
                        PreConditionInfo(patindex).PatternVersion = ParsingArr(m_PatLoop, cst_PatVersion)   '20221121 Load pattern version
                    End If
                    patindex = patindex + 1
                End If
                m_PatLoop = m_PatLoop + 1
            Loop Until (m_PatLoop = SheetMaxRow + 1)
    
            EndIndex = m_PatLoop
    
            For m_PatLoop = 0 To UBound(PreConditionInfo)
            
                Erase asTempPinName
                
                m_start = PreConditionInfo(m_PatLoop).PatCell
                If m_PatLoop = UBound(PreConditionInfo) Then
                    m_end = EndIndex
                Else
                    m_end = PreConditionInfo(m_PatLoop + 1).PatCell
                End If
    
                ''-------------------------------------------------------------------------------------
                ''Get XPins,  Store to "Array data type - Arr_XXXPin" and "String data type - XXXPin"
                ''-------------------------------------------------------------------------------------
                m_XpinIndex = 0
                m_UnusedPinIndex = 0
                m_AllPinIndex = 0
                m_loop = 0
                m_PinStr = vbNullString
                m_AllPinStr = vbNullString
                While (m_loop + m_start <> m_end)
                    sTempPinList = vbNullString
                    If (ParsingArr(m_loop + m_start, cst_XPin) <> "") Then
                        With PreConditionInfo(m_PatLoop)
                            ReDim Preserve .XPins(m_XpinIndex)
                            ReDim Preserve .AllPins(m_AllPinIndex)
                            CellTmp = ParsingArr(m_XpinIndex + m_start, cst_XPin)
                            
                            asTempPinName = Split(CellTmp, ",")
                            
                            For Each vPin In asTempPinName
                                sPinName = CStr(vPin)
                                If dic_nWireLevelName.Exists(sPinName) = False Then sTempPinList = CombineStringList(sTempPinList, sPinName)
                            Next vPin
                            CellTmp = sTempPinList

                            .XPins(m_XpinIndex) = CellTmp
                            .AllPins(m_AllPinIndex) = CellTmp
                        End With
                        If m_XpinIndex = 0 Then
                           m_PinStr = CellTmp
                        Else
                           m_PinStr = m_PinStr + "," + CellTmp
                        End If
                        m_XpinIndex = m_XpinIndex + 1
                        m_AllPinIndex = m_AllPinIndex + 1
                    End If
                    m_loop = m_loop + 1
                Wend
                PreConditionInfo(m_PatLoop).Arr_XPins = Split(m_PinStr, ",")
                m_AllPinStr = m_PinStr
    
                ''-----------------------------------------------------------------------------------------
                ''Get UnusedPins, Store to "Array data type - Arr_XXXPin" and "String data type - XXXPin"
                ''-----------------------------------------------------------------------------------------
                m_PinStr = vbNullString
                m_loop = 0
                While (m_loop + m_start <> m_end)
                    sTempPinList = vbNullString
                    If (ParsingArr(m_loop + m_start, cst_UnusedPin) <> "") Then
                        With PreConditionInfo(m_PatLoop)
                            ReDim Preserve .UnusedPins(m_UnusedPinIndex)
                            ReDim Preserve .AllPins(m_AllPinIndex)
                            CellTmp = ParsingArr(m_UnusedPinIndex + m_start, cst_UnusedPin)
                            
                            asTempPinName = Split(CellTmp, ",")
                            
                            For Each vPin In asTempPinName
                                sPinName = CStr(vPin)
                                If dic_nWireLevelName.Exists(sPinName) = False Then sTempPinList = CombineStringList(sTempPinList, sPinName)
                            Next vPin
                            CellTmp = sTempPinList

                            .UnusedPins(m_UnusedPinIndex) = CellTmp
                            .AllPins(m_AllPinIndex) = CellTmp
                        End With
                        If m_UnusedPinIndex = 0 Then
                            m_PinStr = CellTmp
                        Else
                            m_PinStr = m_PinStr + "," + CellTmp
                        End If
                        m_UnusedPinIndex = m_UnusedPinIndex + 1
                        m_AllPinIndex = m_AllPinIndex + 1
                    End If
                    m_loop = m_loop + 1
                Wend
                PreConditionInfo(m_PatLoop).Arr_UnusedPins = Split(m_PinStr, ",")
                m_AllPinStr = m_AllPinStr + "," + m_PinStr
                PreConditionInfo(m_PatLoop).Arr_AllPins = Split(m_AllPinStr, ",")
                                
                If gl_PreCondition_Dic.Exists(LCase(PreConditionInfo(m_PatLoop).PatternName)) Then
                    gl_PreCondition_Dic.Remove (LCase(PreConditionInfo(m_PatLoop).PatternName))
                End If
                gl_PreCondition_Dic.Add LCase(PreConditionInfo(m_PatLoop).PatternName), m_PatLoop
                
            Next m_PatLoop
            
            gl_ePreConditionError = PreConditionError.PreConditionPass
        End If
        
        
        If gl_ePreConditionError = PreConditionError.SheetNotExist Then
            s_ErrorStatement = sheetName & " doesn't exist in current program. Error!!!"
            Call Print_Error_Message(Warning_Info, LIB_DC, sFuncName, s_ErrorStatement)      '20221205 Tank try use print error
        End If

        ''-------------------------------------------
        ''Get IO pin level from pin level(Levels_IDS)
        ''-------------------------------------------
        ''ex : pins_XpXv
        ''-------------------------------------------
        sheetName = "Levels_IDS"
        
        For i = 0 To UBound(AllSheetNames)
            If LCase(AllSheetNames(i)) Like LCase("*Levels_IDS*") Then
        
                Call GetSheetInfo(sheetName, SheetMaxRow, SheetMaxCol, ParsingArr)
        
                Set tmp = Nothing
                PinIndex = 0
                Do
                    If UCase(ParsingArr(m_loop, PinGroupIndex)) Like "PINS_*" Then
                        If tmp.Exists(ParsingArr(m_loop, PinGroupIndex)) Or ParsingArr(m_loop, PinGroupIndex) = "" Then
                        Else
                            tmp.Add ParsingArr(m_loop, PinGroupIndex), ParsingArr(m_loop, PinGroupIndex)
                            ReDim Preserve PinLevelName(PinIndex)
                            PinLevelName(PinIndex) = ParsingArr(m_loop, PinGroupIndex)
                            PinIndex = PinIndex + 1
                        End If
                    End If
                    m_loop = m_loop + 1
                Loop Until (m_loop = SheetMaxRow + 1)
                
            End If
            
            Levels_IDS_exist = True
            
            
        Next i
        
        If Levels_IDS_exist = False Then
            s_ErrorStatement = sheetName & " doesn't exist in current program. Error!!!"
            Call Print_Error_Message(Warning_Info, LIB_DC, sFuncName, s_ErrorStatement)      '20221205 Tank try use print error
        End If
        
        If Levels_nWire_exist And gl_ePreConditionError = PreConditionError.PreConditionPass And Levels_IDS_exist Then     'If Levels_nWire_exist And gl_isParsePreConditionExist And Levels_IDS_exist Then
            ENG_SweepPin = True
        Else
            ENG_SweepPin = False
        End If
        
        gl_isParsePreConditionDone = True
    End If

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_DC, sFuncName)
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub FindInputPin(PatName As Pattern, _
                        Optional FlagWait As Boolean = False, Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, _
                        Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, _
                        Optional DigSrc_FlowForLoopIntegerName As String, Optional CUS_Str_DigSrcData As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "FindInputPin"
    Dim m_UnusedPinLoop As Long
    Dim m_PatLoop As Long

    Dim m_tmpPatName As String
    Dim m_ChkFlag As Boolean: m_ChkFlag = False
    Dim PinGroup As Variant
    Dim MeasureValue() As New PinListData
    Dim tmp_L() As New PinListData
    Dim tmp_H() As New PinListData
    Dim Force_V(1) As Double
    Dim m_loop As Long
    Dim m_Loop2 As Long
    Dim patindex As Long
    Dim pin As Variant
    Dim force_val As Double
    Dim Tmp_PatName As String: Tmp_PatName = PatName.value
    Const VIL_value = 0
    Const VIH_value = 1

    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "Start to find input pins"
    TheExec.Datalog.WriteComment ""
    
    If UCase(Tmp_PatName) Like "IDS_PP_*" Then Tmp_PatName = Replace(UCase(Tmp_PatName), "IDS_PP_", "PP_")

    ''----------------------------------------------
    ''Set Init status of UnusedPin to Initoff first.
    ''----------------------------------------------
    SetPinInitStatus Tmp_PatName, chInitoff, patindex
    If FindIDSPattern = False Then Exit Sub

''    For m_PatLoop = 0 To UBound(PreConditionInfo)
''        m_UnusedPinLoop = 0
''        m_tmpPatName = PreConditionInfo(m_PatLoop).PatternName
''        If UCase(patname.Value) = UCase(m_tmpPatName) Then
''            While m_UnusedPinLoop <= UBound(PreConditionInfo(m_PatLoop).UnusedPins)
''                thehdw.Digital.Pins(PreConditionInfo(m_PatLoop).UnusedPins(m_UnusedPinLoop)).InitState = chInitoff
''                m_UnusedPinLoop = m_UnusedPinLoop + 1
''            Wend
''            PatIndex = m_PatLoop
''            m_ChkFlag = True
''            Theexec.Datalog.WriteComment "< UnusedPin > PreCondition Setting : Initoff."
''        End If
''    Next m_PatLoop

    ''--------------------------------------------------------------------
    ''exit function if can't find the pattern name from precondition table.
    ''--------------------------------------------------------------------
''    If Not m_ChkFlag Then
''        Theexec.Datalog.WriteComment "Can't find the pattern from < PreCondition > Table."
''        Exit Sub
''    End If

    ''---------------------------------------------------------------
    ''Clear all CPU flag & disconnet don't care pins and unused pins.
    ''---------------------------------------------------------------
    With thehdw.Digital
        .Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
        .ApplyLevelsTiming True, True, True, tlPowered
        .Patgen.TimeOut = 10
        For m_loop = 0 To UBound(PreConditionInfo(patindex).XPins)
            .Pins(PreConditionInfo(patindex).XPins(m_loop)).Disconnect
        Next
        For m_loop = 0 To UBound(PreConditionInfo(patindex).UnusedPins)
            .Pins(PreConditionInfo(patindex).UnusedPins(m_loop)).Disconnect
        Next
    End With

    ''-------------------------------------
    ''Get VIH and VIL from pin level sheet.
    ''-------------------------------------
    GetVihVilValue patindex

''    Dim Pin_Ary() As String
''    Dim Pin_cnt As Long
''    Dim VIH_val As Double
''    Dim VIL_val As Double
''
''    dic_forceV.RemoveAll
''    dic_VIL.RemoveAll
''
''    For m_Loop = 0 To UBound(PinLevelName)
''        Theexec.DataManager.DecomposePinList PinLevelName(m_Loop), Pin_Ary, Pin_cnt
''        ''hard code for T-ell
''        If PinLevelName(m_Loop) = "Pins_1p1v" Then
''            VIH_val = 0.99
''            VIL_val = 0.11
''        ElseIf PinLevelName(m_Loop) = "Pins_0p6v" Then
''            VIH_val = 0.54
''            VIL_val = 0.06
''        Else
''            VIH_val = thehdw.Digital.Pins(PinLevelName(m_Loop)).Levels.Value(chVih)
''            VIL_val = thehdw.Digital.Pins(PinLevelName(m_Loop)).Levels.Value(chVil)
''        End If
''        For m_Loop2 = 0 To Pin_cnt - 1
''            If dic_forceV.Exists(Pin_Ary(m_Loop2)) Then
''            Else
''                Theexec.Datalog.WriteComment Pin_Ary(m_Loop2) & " - VIH - " & VIH_val
''                dic_forceV.Add Pin_Ary(m_Loop2), VIH_val
''                dic_VIL.Add Pin_Ary(m_Loop2), VIL_val
''            End If
''        Next
''    Next

    ''--------------------
    ''run the IDS pattern.
    ''--------------------
    Dim InDSPWave As New DSPWave
    Dim patCnt As Long
    Dim rtnPatNames() As String

    If PatName <> "" Then
        Call thehdw.Patterns(PatName).Load
        If DigSrc_pin <> "" Then
            rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(PatName, patCnt)
            Call GeneralDigSrcSetting(rtnPatNames(0), DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, _
                DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPWave)
        End If
        If FlagWait = True Then
            Call thehdw.Patterns(PatName).start
            Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)  'Meas during CPUA loop
        Else
            Call thehdw.Patterns(PatName).test(pfAlways, 0, tlResultModeDomain)
            thehdw.Digital.Patgen.HaltWait
        End If
    Else
    End If
    
    ''------------------------------
    ''1.Force VIL / VIH for each pin
    ''2.Measure current
    ''------------------------------
    Dim tmpvih0 As String
    'wait for verify
    'Call IOpinsForceVMeasureI(PatIndex, tmp_L, tmp_H, tmpvih0)
    ReDim tmp_L(UBound(PreConditionInfo(patindex).AllPins))
    ReDim tmp_H(UBound(PreConditionInfo(patindex).AllPins))
    ReDim MeasureValue(UBound(PreConditionInfo(patindex).AllPins))

    For m_loop = VIL_value To VIH_value
        If m_loop = VIL_value Then
            'force_val = 0
            TheExec.Datalog.WriteComment "--------------------------"
            TheExec.Datalog.WriteComment " force VIL to measure I"
            TheExec.Datalog.WriteComment "--------------------------"
        Else
             TheExec.Datalog.WriteComment "--------------------------"
            TheExec.Datalog.WriteComment " force VIH to measure I"
            TheExec.Datalog.WriteComment "--------------------------"
        End If

        For Each pin In PreConditionInfo(patindex).Arr_AllPins
            If dic_IgnorePin.Exists(pin) Then
                tmpvih0 = tmpvih0 + IIf(tmpvih0 <> "", ";", "") + pin
                GoTo skip1
            End If
            If m_loop = VIH_value Then
                force_val = dic_forceV.item(pin)
                 'If force_val = 0 Then tmpvih0 = tmpvih0 + IIf(tmpvih0 <> "", ";", "") + pin
            ElseIf m_loop = VIL_value Then
                force_val = dic_VIL.item(pin)
            End If
            If Enable_SweepPinPrintOut Then TheExec.Datalog.WriteComment pin & " - force V value - " & force_val
            With thehdw.PPMU.Pins(pin)
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                .Connect
                .Gate = tlOn
                .ForceV force_val, pc_Def_PPMU_Max_InitialValue_FI_Range
            End With
skip1:
        Next pin

        m_Loop2 = 0
        For Each PinGroup In PreConditionInfo(patindex).AllPins
            thehdw.Wait (100 * us)
            MeasureValue(m_Loop2) = thehdw.PPMU.Pins(PinGroup).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            m_Loop2 = m_Loop2 + 1
            With thehdw.PPMU.Pins(PinGroup)
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                .Disconnect
                .Gate = tlOff
            End With
            'TheHdw.Digital.Pins(pingroup).Connect
        Next PinGroup

        For m_Loop2 = 0 To UBound(PreConditionInfo(patindex).AllPins)
            If m_loop = 0 Then
                tmp_L(m_Loop2) = MeasureValue(m_Loop2)
            Else
                tmp_H(m_Loop2) = MeasureValue(m_Loop2)
            End If
        Next
    Next m_loop

    thehdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
    thehdw.Digital.Patgen.HaltWait

    Dim p As Long
    Dim tmpstr As String
    Dim m_site As Variant

'    If Theexec.TesterMode = testModeOffline Then
'        For m_Loop = 0 To 2
'            For p = 0 To tmp_L(m_Loop).Pins.Count - 1
'                For Each m_site In Theexec.sites
'                    tmp_L(m_Loop).Pins(p).Value(m_site) = 0.00003 + Rnd() * 0.00005
'                    tmp_H(m_Loop).Pins(p).Value(m_site) = 0.00003 + Rnd() * 0.00005
'                Next m_site
'            Next p
'        Next m_Loop
'    End If

    ''----------------------------------------------------
    ''Compare forcing VIL current and forcing VIH current.
    ''----------------------------------------------------
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "Sort out the input pin if the measured value of Vt =0  is lower than 50uA and the measured value of Vt =VIH is lower than 50uA."

    'wait for verify
    'CompareCurrentForVilVih tmp_L, tmp_H, tmpvih0, PatIndex, patname.Value

    Dim value1  As Double
    Dim value2 As Double
    Dim dictcheck As Boolean
    Dim pinnamecheck As Boolean

    m_Loop2 = 0
    For m_loop = 0 To UBound(PreConditionInfo(patindex).AllPins)
        For p = 0 To tmp_L(m_loop).Pins.Count - 1
            'pinnamecheck = False
            If dic_IgnorePin.Exists(tmp_L(m_loop).Pins(p).name) Then
                TheExec.Datalog.WriteComment tmp_L(m_loop).Pins(p).name + " - the pin can't find the VIL/VIH from level sheet!!!"
                GoTo skip2
            Else
                TheExec.Datalog.WriteComment tmp_L(m_loop).Pins(p).name
            End If
            For Each m_site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "tmp_L value -> " & Format(tmp_L(m_loop).Pins(p).value(m_site) * 1000, "0.000000") & "mA"
                TheExec.Datalog.WriteComment "tmp_H value -> " & Format(tmp_H(m_loop).Pins(p).value(m_site) * 1000, "0.000000") & "mA"
                thehdw.Wait 0.01
                value1 = Abs(tmp_L(m_loop).Pins(p).value(m_site))
                value2 = Abs(tmp_H(m_loop).Pins(p).value(m_site))

                'If Not (UCase(tmp_L(m_Loop).Pins(p).Name) Like "*_PA") Then pinnamecheck = True
                    If value1 < 0.00005 And value2 < 0.00005 Then
                    'If value1 < 0.00005 And value2 < 0.00005 And pinnamecheck Then
                        If InStr(1, tmpvih0, UCase(tmp_L(m_loop).Pins(p).name), vbBinaryCompare) = 0 Then
                        
                            ReDim Preserve PreConditionInfo(patindex).Arr_InputPins(m_Loop2)
                            PreConditionInfo(patindex).Arr_InputPins(m_Loop2) = tmp_L(m_loop).Pins(p).name
                            If m_Loop2 = 0 Then
                                tmpstr = tmp_L(m_loop).Pins(p).name
                            Else
                                tmpstr = tmpstr + "," + tmp_L(m_loop).Pins(p).name
                            End If
                            TheExec.Datalog.WriteComment "------------------------------------------------------"
                            TheExec.Datalog.WriteComment tmp_L(m_loop).Pins(p).name & " -> InputPin"
                            TheExec.Datalog.WriteComment "------------------------------------------------------"
                            m_Loop2 = m_Loop2 + 1
                            Exit For
                        End If
                    End If
skip2:
            Next m_site
        Next p
    Next m_loop

    TheExec.Datalog.WriteComment "InputPins - " & tmpstr
    PreConditionInfo(patindex).InputPins = tmpstr
    CurrentPatIndex = patindex
    If dic_mapCritcalPin.Exists(Tmp_PatName) Then
    Else
        dic_mapCritcalPin.Add Tmp_PatName, patindex
    End If
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "FindInputPin") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub

    
    

Public Sub IDS_PatternRun(PatName As Pattern, _
                          Optional FlagWait As Boolean = False, Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, _
                          Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, _
                          Optional DigSrc_FlowForLoopIntegerName As String, Optional CUS_Str_DigSrcData As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim InDSPWave As New DSPWave
    Dim patCnt As Long
    Dim rtnPatNames() As String
    Dim m_loop As Long

    With thehdw.Digital
        .Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
        .ApplyLevelsTiming True, True, True, tlPowered
        .Patgen.TimeOut = 10
    End With

    If PatName <> "" Then
        Call thehdw.Patterns(PatName).Load
        If DigSrc_pin <> "" Then
            rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(PatName, patCnt)
            Call GeneralDigSrcSetting(rtnPatNames(0), DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, _
                DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPWave)
        End If
        If FlagWait = True Then
            Call thehdw.Patterns(PatName).start
            Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)  'Meas during CPUA loop
        Else
            Call thehdw.Patterns(PatName).test(pfAlways, 0, tlResultModeDomain)
            thehdw.Digital.Patgen.HaltWait
        End If
    Else
    End If
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "IDS_PatternRun") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub

Public Sub SweepPinInPattern(PatName As Pattern, FlagWait As Boolean, DCVS_Power_Pin As String, DCVS_OtherPower_Pin As String, _
                              repeat_count As Long, FlowLimitForInitIRange As Boolean, _
                              Optional DisableClock As Boolean = False, Optional Search_Step As String, _
                              Optional DisableClockPortName As String, Optional DisconnectClock As Boolean = True, _
                              Optional debug_print_pins As String, Optional AutoRange_Pin As String, Optional DigSrc_pin As PinList, _
                              Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, _
                              Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String, Optional CUS_Str_DigSrcData As String, _
                              Optional Interpose_Meas_before As String, Optional Interpose_Meas_after As String, Optional DisableFRCPinName As String, Optional FRC_RelayPin As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim patindex As Long: patindex = CurrentPatIndex
    Dim m_loop As Long
    Dim Vt_val As Double
    Dim Val_Ratio As New SiteDouble
    Dim SD_vt_val As New SiteDouble
    Dim Vt_Switch As Long
    Dim All_Power_data_VIL As New PinListData
    Dim All_Power_data_VIH As New PinListData
    Dim All_Power_data_Half As New PinListData
    Dim All_Power_data_Sweep() As New PinListData
    Dim p As Long
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim Sweep_Flag As Boolean: Sweep_Flag = False
    Dim m_site As Variant
    Dim SweepPinIndex As Long
    Dim CriticalStr As String
    Dim m_InputPin() As String
    Dim i As Long
    Dim tmpstr As String
    Dim InputIndex As Long
    Dim counter As Long
    Dim All_Power As String
    Dim ratio_count As Integer
    Dim critical_flag As Boolean
    Dim PrintOutTmp As New SiteVariant
    Dim RatioStr As String
    Dim Vt_Status As Integer
    Dim Vt_Status_Str As String

    Const Vt_All_VIL = 0
    Const Vt_All_VIH = 2
    Const Vt_All_Half = 1
    Const Vt_Sweep_Half = 3
    Const End_Loop = 4
    Const MidVt = 0
    Const HighVt = 1
    
    If FindIDSPattern = False Then Exit Sub

    ReDim All_Power_data_Sweep(UBound(PreConditionInfo(patindex).Arr_InputPins)) As New PinListData

    If DCVS_Power_Pin <> "" Then
        If (UCase(DCVS_Power_Pin) Like "*CP*=*" Or UCase(DCVS_Power_Pin) Like "*FT*=*") Then DCVS_Power_Pin = Select_MeasPin(DCVS_Power_Pin, UCase(currentJobName))
        If (UCase(AutoRange_Pin) Like "*CP*=*" Or UCase(AutoRange_Pin) Like "*FT*=*") Then AutoRange_Pin = Select_MeasPin(AutoRange_Pin, UCase(currentJobName))
    End If

        tmpstr = vbNullString
        InputIndex = 0
        counter = 0
        For i = 0 To UBound(PreConditionInfo(patindex).Arr_InputPins)
            tmpstr = IIf(counter <> 0, tmpstr + "," + PreConditionInfo(patindex).Arr_InputPins(i), PreConditionInfo(patindex).Arr_InputPins(i))
            counter = i
            If Len(tmpstr) >= 8000 Or i = UBound(PreConditionInfo(patindex).Arr_InputPins) Then
                tmpstr = Replace(tmpstr, "," + PreConditionInfo(patindex).Arr_InputPins(i), "")
                ReDim Preserve m_InputPin(InputIndex)
                m_InputPin(InputIndex) = tmpstr
                tmpstr = PreConditionInfo(patindex).Arr_InputPins(i)
                counter = 0
                InputIndex = InputIndex + 1
            End If
        Next

    All_Power = DCVS_Power_Pin + "," + DCVS_OtherPower_Pin
    If right(All_Power, 1) = "," Then All_Power = left(All_Power, Len(All_Power) - 1)
    TheExec.DataManager.DecomposePinList All_Power, Pin_Ary, Pin_Cnt

    AutoRange_Pin = All_Power

    SweepPinIndex = 0
    Vt_Switch = Vt_All_VIL
    Vt_Status = MidVt
    Do
        Call IDS_PatternRun(PatName, FlagWait, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
                            DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData)
        For m_loop = 0 To UBound(PreConditionInfo(patindex).Arr_InputPins)
            thehdw.Digital.Pins(PreConditionInfo(patindex).Arr_InputPins(m_loop)).Levels.DriverMode = tlDriverModeVt
        Next
        thehdw.Wait 0.01

        Select Case Vt_Switch
            Case Vt_All_VIL

                For m_loop = 0 To UBound(PreConditionInfo(patindex).Arr_InputPins)
                   thehdw.Digital.Pins(PreConditionInfo(patindex).Arr_InputPins(m_loop)).Levels.value(chVt) = 0
                Next

                TheExec.Datalog.WriteComment String(63, "-")
                TheExec.Datalog.WriteComment "< All Input I/O Pins Set Vt = VIL >"
                TheExec.Datalog.WriteComment String(63, "-")


                MeasureIDS All_Power_data_VIL, PatName, FlagWait, DCVS_Power_Pin, DCVS_OtherPower_Pin, _
                              repeat_count, FlowLimitForInitIRange, _
                              DisableClock, Search_Step, _
                              DisableClockPortName, DisconnectClock, _
                              debug_print_pins, AutoRange_Pin, DigSrc_pin, _
                              DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, _
                              DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, _
                              Interpose_Meas_before, Interpose_Meas_after, DisableFRCPinName, FRC_RelayPin
                
                If SkipAllPinToHalf Then
                    Vt_Switch = Vt_Sweep_Half   ''Jump to sweep critical
                Else
                    Vt_Switch = Vt_All_Half     ''Jump to all pin set to vt = (VIH+VIL)/2
                End If
                
            Case Vt_All_Half
                TheExec.Datalog.WriteComment String(63, "-")
                TheExec.Datalog.WriteComment "< All Input I/O Pins Set Vt = (VIL+VIH)/2 >"
                TheExec.Datalog.WriteComment String(63, "-")

                For m_loop = 0 To UBound(PreConditionInfo(patindex).Arr_InputPins)
                    Vt_val = (dic_forceV.item(UCase(PreConditionInfo(patindex).Arr_InputPins(m_loop))) + _
                                dic_VIL.item(UCase(PreConditionInfo(patindex).Arr_InputPins(m_loop)))) / 2
                   thehdw.Digital.Pins(PreConditionInfo(patindex).Arr_InputPins(m_loop)).Levels.value(chVt) = Vt_val
                Next
                


                MeasureIDS All_Power_data_Half, PatName, FlagWait, DCVS_Power_Pin, DCVS_OtherPower_Pin, _
                              repeat_count, FlowLimitForInitIRange, _
                              DisableClock, Search_Step, _
                              DisableClockPortName, DisconnectClock, _
                              debug_print_pins, AutoRange_Pin, DigSrc_pin, _
                              DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, _
                              DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, _
                              Interpose_Meas_before, Interpose_Meas_after

                Vt_Switch = End_Loop
                For p = 0 To Pin_Cnt - 1
                    If TheExec.DataManager.ChannelType(Pin_Ary(p)) = "N/C" Then GoTo skip_calc2
                    TheExec.Datalog.WriteComment All_Power_data_Half.Pins(Pin_Ary(p)).name

                    For Each m_site In TheExec.sites.Active
                        If All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) = 0 Then
                            TheExec.Datalog.WriteComment "All_Power_data_VIL -" & Pin_Ary(p) & "IDS value = 0 "
                            GoTo skip_calc
                        End If
                        TheExec.Datalog.WriteComment "All_Power_data_Half - " & Format(All_Power_data_Half.Pins(Pin_Ary(p)).value(m_site) * 1000, "0.000000") & "mA"
                        TheExec.Datalog.WriteComment "All_Power_data_VIL - " & Format(All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) * 1000, "0.000000") & "mA"
                        'If All_Power_data_VIL.Pins(Pin_Ary(p)).Value(m_site) = 0 Then Stop
                        'Vt_val = Abs(All_Power_data_Half.Pins(Pin_Ary(p)).Value(m_site) - All_Power_data_VIL.Pins(Pin_Ary(p)).Value(m_site)) / All_Power_data_VIL.Pins(Pin_Ary(p)).Value(m_site) * 100
                        Vt_val = Abs(All_Power_data_Half.Pins(Pin_Ary(p)).value(m_site) / All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site))
                        TheExec.Datalog.WriteComment "Vt_val (%) - " & Format(Vt_val, "0.000")
                        If Vt_val > Threshold_EnableSweepPin Then ''Threshold_EnableSweepPin = 8
                            Vt_Switch = Vt_Sweep_Half
                            TheExec.Datalog.WriteComment "Need to sweep critical pin one by one in pattern!!!"
                        End If
skip_calc:
                    Next
skip_calc2:
                Next p
            Case Vt_Sweep_Half
                ''Initialize Vt of all input pin to VIL
                If SweepPinIndex = 0 And Vt_Status < 2 Then
                    With TheExec.Datalog
                        .WriteComment String(63, "-")
                        If Vt_Status = MidVt Then
                            .WriteComment "< Sweep All Input I/O Pins In Pattern and Set Vt = (VIL+VIH)/2 >"
                            Vt_Status_Str = "(VIL+VIH)/2"
                        ElseIf Vt_Status = HighVt Then
                            .WriteComment "< Sweep All Input I/O Pins In Pattern and Set Vt = VIH >"
                            ReDim All_Power_data_Sweep(UBound(PreConditionInfo(patindex).Arr_InputPins)) As New PinListData
                            Vt_Status_Str = "VIH"
                        End If
                        .WriteComment String(63, "-")
                    End With
                End If

                For m_loop = 0 To UBound(PreConditionInfo(patindex).Arr_InputPins)
                    thehdw.Digital.Pins(PreConditionInfo(patindex).Arr_InputPins(m_loop)).Levels.value(chVt) = 0
                Next

                If Vt_Status = MidVt Then
                    Vt_val = (dic_forceV.item(UCase(PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex))) + _
                                dic_VIL.item(UCase(PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex)))) / 2
                ElseIf Vt_Status = HighVt Then
                    Vt_val = dic_forceV.item(UCase(PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex)))
                End If
                thehdw.Digital.Pins(PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex)).Levels.value(chVt) = Vt_val
                TheExec.Datalog.WriteComment ""
                TheExec.Datalog.WriteComment "-----Sweep pin" & SweepPinIndex & " - " & PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex) & _
                    " Vt = " & CStr(Vt_val) & "V -----"


                ''halt pattern (this's for HDC proj request - halt before measure)
''                TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
''                TheHdw.Digital.Patgen.HaltWait

                MeasureIDS All_Power_data_Sweep(SweepPinIndex), PatName, FlagWait, DCVS_Power_Pin, DCVS_OtherPower_Pin, _
                                repeat_count, FlowLimitForInitIRange, _
                                DisableClock, Search_Step, _
                                DisableClockPortName, DisconnectClock, _
                                debug_print_pins, AutoRange_Pin, DigSrc_pin, _
                                DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, _
                                DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, _
                                Interpose_Meas_before, Interpose_Meas_after

                TheExec.Datalog.WriteComment ""

                critical_flag = False
                For p = 0 To Pin_Cnt - 1
                    ratio_count = 0
                    If TheExec.DataManager.ChannelType(Pin_Ary(p)) = "N/C" Then GoTo skip_loop
                    RatioStr = vbNullString
                    PrintOutTmp = vbNullString
                    For Each m_site In TheExec.sites
                         ''---------------------------------------------------------------
                         ''caseA : all I/O pin vt = 0            - IDS threshold: 0.1 uA
                         ''caseB : sweep I/ pin vt = (VIL+VIH)/2 - IDS threshold: 1   uA
                         ''---------------------------------------------------------------

'                        If All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).Value(m_site) > 0.000001 And _
'                            All_Power_data_VIL.Pins(Pin_Ary(p)).Value(m_site) > 0.0000001 And _
'                            All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).Value(m_site) > All_Power_data_VIL.Pins(Pin_Ary(p)).Value(m_site) Then
                         If All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) > Threshold_VT_Half And _
                            Abs(All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site)) > 0 And _
                            All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) > All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) Then

                            If All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) < Threshold_VT_Vil Then
                                ratio_count = ratio_count + 1
                                RatioStr = " > " + CStr(Threshold_CriticalPin)
                                PrintOutTmp = "<Error>"
                            Else
                                Vt_val = Abs(All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) - All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site)) / All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) * 100
                                If All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) > All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) Then
                                    Val_Ratio(m_site) = Ceiling(All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) / All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site))
                                Else
                                    Val_Ratio(m_site) = Ceiling(All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) / All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site))
                                End If

                                If Val_Ratio(m_site) > Threshold_CriticalPin Then
                                    ratio_count = ratio_count + 1
                                    PrintOutTmp = "<Error>"
                                End If
                                RatioStr = CStr(Val_Ratio(m_site))
                            End If
                            PrintOutTmp = PrintOutTmp + _
                              "IDS_Check, site " + CStr(m_site) + " ,I/O PinName " + PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex) + _
                              ",Power Pin " + All_Power_data_VIL.Pins(Pin_Ary(p)).name + _
                              ", VIL " + Format((All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) * 1000), "0.000000") & " mA" + ", " + Vt_Status_Str + " " + Format((All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) * 1000), "0.000000") & " mA" + _
                              ", Ratio " + RatioStr
                        Else
                            If All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) < Threshold_VT_Half Then
                                PrintOutTmp = Vt_Status_Str + " IDS < 1uA"
                            ElseIf All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) < All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) Then
                                 PrintOutTmp = Vt_Status_Str + " IDS < VIL IDS"
                            End If
                            PrintOutTmp = "IDS_Check, site " + CStr(m_site) + " ,I/O PinName " + PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex) + _
                              ",Power Pin " + All_Power_data_VIL.Pins(Pin_Ary(p)).name + _
                              ", VIL " + Format((All_Power_data_VIL.Pins(Pin_Ary(p)).value(m_site) * 1000), "0.000000") & " mA" + ", " + Vt_Status_Str + " " + Format((All_Power_data_Sweep(SweepPinIndex).Pins(Pin_Ary(p)).value(m_site) * 1000), "0.000000") & " mA" + _
                              ", " + PrintOutTmp
                        End If
                    Next
                    If ratio_count = TheExec.sites.Active.Count Then
                       critical_flag = True
                    Else
                        For Each m_site In TheExec.sites
                            If Val_Ratio(m_site) > Threshold_CriticalPin Then
                                PrintOutTmp = "<Warning> higher than 5 times! " & PrintOutTmp
                            Else
                            End If
                        Next
                    End If
                    For Each m_site In TheExec.sites.Active
                        TheExec.Datalog.WriteComment PrintOutTmp
                    Next
skip_loop:
                Next p
                If critical_flag = True Then
                    If CriticalStr = "" Then
                        CriticalStr = PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex)
                        Else
                            CriticalStr = CriticalStr + "," + PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex)
                        End If
                        With TheExec.Datalog
                            .WriteComment String(60, "*")
                            .WriteComment PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex) & " is critical pin."
                            .WriteComment String(60, "*")
                        End With
                End If
                thehdw.Digital.Pins(PreConditionInfo(patindex).Arr_InputPins(SweepPinIndex)).Levels.value(chVt) = 0
                If SweepPinIndex = UBound(PreConditionInfo(patindex).Arr_InputPins) Then
                    If Vt_Status = MidVt Then
                        Vt_Status = HighVt
                        SweepPinIndex = -1
                    Else
                        Vt_Switch = End_Loop
                    End If
                End If
                SweepPinIndex = SweepPinIndex + 1
        End Select
    Loop Until (Vt_Switch = End_Loop)

    PreConditionInfo(patindex).Arr_CriticalPins = Split(CriticalStr, ",")
    PreConditionInfo(patindex).CriticalPins = CriticalStr
    TheExec.Datalog.WriteComment String(100, "-")
    TheExec.Datalog.WriteComment "CriticalPins - " & PreConditionInfo(patindex).CriticalPins
    TheExec.Datalog.WriteComment String(100, "-")
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "SweepPinInPattern") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Public Sub SetCriticalPins_InitLo(Optional CriticalPins As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "---------------------------------------"
    TheExec.Datalog.WriteComment "CriticalPin - " & CriticalPins
    TheExec.Datalog.WriteComment "---------------------------------------"
    If CriticalPins = "" Then Exit Sub

    thehdw.Digital.Pins(CriticalPins).initState = chInitLo

    With TheExec.Datalog
        .WriteComment String(60, "-")
        .WriteComment "< CriticalPin > Set InitState = InitLo"
        .WriteComment String(60, "-")
    End With
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "SetCriticalPins_InitLo") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Public Sub SetCriticalPins_RelayOff(Optional CriticalPins As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
 
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "---------------------------------------"
    TheExec.Datalog.WriteComment "CriticalPin - " & CriticalPins
    TheExec.Datalog.WriteComment "---------------------------------------"
    If CriticalPins = "" Then Exit Sub

    thehdw.Digital.Pins(CriticalPins).Disconnect

    With TheExec.Datalog
        .WriteComment String(60, "-")
        .WriteComment "< CriticalPin > Disconnect."
        .WriteComment String(60, "-")
    End With
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "SetCriticalPins_RelayOff") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Public Sub SetCriticalPins_VtMode(Optional CriticalPins As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "---------------------------------------"
    TheExec.Datalog.WriteComment "CriticalPin - " & CriticalPins
    TheExec.Datalog.WriteComment "---------------------------------------"
    If CriticalPins = "" Then Exit Sub
    
    With thehdw.Digital.Pins(CriticalPins)
        .Levels.DriverMode = tlDriverModeVt
        .Levels.value(chVt) = 0
    End With
    
    With TheExec.Datalog
        .WriteComment String(60, "-")
        .WriteComment "< CriticalPin > Set to Vt mode and Vt = 0."
        .WriteComment String(60, "-")
    End With
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "SetCriticalPins_VtMode") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub

Public Sub SetPinInitStatus(PatName As String, initStatus As Long, patindex As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim m_PatLoop As Long
    Dim m_UnusedPinLoop As Long
    Dim m_tmpPatName As String
    'Dim m_ChkFlag As Boolean: m_ChkFlag = False
    'Dim PatIndex As Long
    
    FindIDSPattern = False
    
    For m_PatLoop = 0 To UBound(PreConditionInfo)
        m_UnusedPinLoop = 0
        m_tmpPatName = PreConditionInfo(m_PatLoop).PatternName
        If UCase(PatName) = UCase(m_tmpPatName) Then
            While m_UnusedPinLoop <= UBound(PreConditionInfo(m_PatLoop).UnusedPins)
                thehdw.Digital.Pins(PreConditionInfo(m_PatLoop).UnusedPins(m_UnusedPinLoop)).initState = chInitoff
                m_UnusedPinLoop = m_UnusedPinLoop + 1
            Wend
            patindex = m_PatLoop
            FindIDSPattern = True
            TheExec.Datalog.WriteComment "< UnusedPin > PreCondition Setting : Initoff."
        End If
    Next m_PatLoop

    ''--------------------------------------------------------------------
    ''exit function if can't find the pattern name from precondition table.
    ''--------------------------------------------------------------------
    If Not FindIDSPattern Then
        TheExec.Datalog.WriteComment "Can't find the pattern from < PreCondition > Table."
        'Exit Sub
    End If

Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "SetPinInitStatus") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub

Public Sub GetVihVilValue(patindex As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String: funcName = "GetVihVilValue"

    Dim m_loop As Long
    Dim m_Loop2 As Long
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim VIH_val As Double
    Dim VIL_val As Double
    Dim pin As Variant

    dic_forceV.RemoveAll
    dic_VIL.RemoveAll

    For m_loop = 0 To UBound(PinLevelName)
        TheExec.DataManager.DecomposePinList PinLevelName(m_loop), Pin_Ary, Pin_Cnt
        ''hard code for T-ell
        If TheExec.TesterMode = testModeOffline Then
            If PinLevelName(m_loop) = "Pins_1p8v" Then
                VIH_val = 1.62
                VIL_val = 0.18
            ElseIf PinLevelName(m_loop) = "Pins_1p2v_IO" Then
                VIH_val = 1.08
                VIL_val = 0.12
            ElseIf PinLevelName(m_loop) = "Pins_1p1v" Then
                VIH_val = 0.99
                VIL_val = 0.11
            ElseIf PinLevelName(m_loop) = "Pins_0p6v" Then
                VIH_val = 0.54
                VIL_val = 0.06
            Else
                VIH_val = thehdw.Digital.Pins(Pin_Ary(0)).Levels.value(chVih)
                VIL_val = thehdw.Digital.Pins(Pin_Ary(0)).Levels.value(chVil)
            End If
        Else
            If TheExec.DataManager.ChannelType(Pin_Ary(0)) <> "N/C" Then
                VIH_val = thehdw.Digital.Pins(Pin_Ary(0)).Levels.value(chVih)
                VIL_val = thehdw.Digital.Pins(Pin_Ary(0)).Levels.value(chVil)
            End If
        End If
        For m_Loop2 = 0 To Pin_Cnt - 1
            If TheExec.DataManager.ChannelType(Pin_Ary(m_Loop2)) <> "N/C" Then
                If dic_forceV.Exists(Pin_Ary(m_Loop2)) Then
                Else
                    If Enable_SweepPinPrintOut Then
                        TheExec.Datalog.WriteComment Pin_Ary(m_Loop2) & " - VIH - " & VIH_val
                        TheExec.Datalog.WriteComment Pin_Ary(m_Loop2) & " - VIL - " & VIL_val
                    End If
                    dic_forceV.Add Pin_Ary(m_Loop2), VIH_val
                    dic_VIL.Add Pin_Ary(m_Loop2), VIL_val
                End If
            End If
        Next
    Next

    ''-------------------------------------------
    ''Add ignored I/O pin(not exist in Level_IDS)
    ''-------------------------------------------
    ''Add ignored I/O pin for VDD***SENSE pin
    ''-------------------------------------------

    m_loop = 0
    For Each pin In PreConditionInfo(patindex).Arr_AllPins
        If dic_forceV.Exists(pin) Then
            If UCase(pin) Like "VDD*SENSE*" Then
                If dic_IgnorePin.Exists(pin) Then
                Else
                    dic_IgnorePin.Add pin, m_loop
                    m_loop = m_loop + 1
                End If
                If Enable_SweepPinPrintOut Then TheExec.Datalog.WriteComment "Ignore pin - " & pin & " is a power sense pin!!!"
            End If
        Else
            If dic_IgnorePin.Exists(pin) Then
            Else
                dic_IgnorePin.Add pin, m_loop
                m_loop = m_loop + 1
            End If
            If Enable_SweepPinPrintOut Then TheExec.Datalog.WriteComment "Ignore pin - " & pin + " can't find the VIH and VIL from Level_IDS ,it exists in Level_Nwire ,or it's N/C pin!!!"
        End If
    Next pin

Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "GetVihVilValue") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Public Sub MeasureIDS(All_Power_data As PinListData, PatName As Pattern, FlagWait As Boolean, DCVS_Power_Pin As String, DCVS_OtherPower_Pin As String, _
                              repeat_count As Long, FlowLimitForInitIRange As Boolean, _
                              Optional DisableClock As Boolean = False, Optional Search_Step As String, _
                              Optional DisableClockPortName As String, Optional DisconnectClock As Boolean = True, _
                              Optional debug_print_pins As String, Optional AutoRange_Pin As String, Optional DigSrc_pin As PinList, _
                              Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, _
                              Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String, Optional CUS_Str_DigSrcData As String, _
                              Optional Interpose_Meas_before As String, Optional Interpose_Meas_after As String, Optional DisableFRCPinName As String, Optional FRC_RelayPin As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If DisableClock = True Then
        If DisableFRCPinName <> "" And FRC_RelayPin <> "" Then
            Call Disable_FRC_Pins(DisableFRCPinName, FRC_RelayPin, DisableClock)
        Else
            Call Disable_FRC(DisableClockPortName, DisconnectClock)
        End If
        Wait 0.005
    End If

    If Interpose_Meas_before <> "" Then
        Call SetForceCondition(Interpose_Meas_before & ";STOREPREMEAS")
    End If

    If DCVS_Power_Pin <> "" Then
        If (UCase(DCVS_Power_Pin) Like "*CP*=*" Or UCase(DCVS_Power_Pin) Like "*FT*=*") Then DCVS_Power_Pin = Select_MeasPin(DCVS_Power_Pin, UCase(currentJobName))
        If (UCase(AutoRange_Pin) Like "*CP*=*" Or UCase(AutoRange_Pin) Like "*FT*=*") Then AutoRange_Pin = Select_MeasPin(AutoRange_Pin, UCase(currentJobName))
        DCVS_IDS_main_auto_range_and_measure DCVS_Power_Pin, DCVS_OtherPower_Pin, All_Power_data, repeat_count, FlowLimitForInitIRange, Search_Step, debug_print_pins, AutoRange_Pin
    End If


    If DisableClock = True Then
        If DisableFRCPinName <> "" And FRC_RelayPin <> "" Then
            Call Enable_FRC_Pins(DisableFRCPinName, FRC_RelayPin, DisableClock)
        Else
            Call Enable_FRC(DisableClockPortName, DisableClock)
        End If
    
    End If
                '-------------------------------------------
    thehdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
    thehdw.Digital.Patgen.HaltWait

    If Interpose_Meas_after <> "" Then
        Call SetForceCondition(Interpose_Meas_after)
    ElseIf Interpose_Meas_before <> "" Then
        Call SetForceCondition("RESTOREPREMEAS")
    Else

    End If
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "MeasureIDS") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub

' [20231228][T-Bra][Tank] Modify input set CurrentRange by each pin
Public Function FixedIRange_Program_DCVS(powerPin As String, ByRef pin_info As AutoRange_Info, d_IncludeCurrent As Double, ByRef WaitTime As Double, ByRef SattleTime As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
                                    
    Dim funcName As String:: funcName = "FixedIRange_Program_DCVS"
    
    Dim i As Long
    Dim k As Long
    Dim d_FixedCurrent As Double
    Dim n_Power_ary_Index As Long
    
    d_FixedCurrent = d_IncludeCurrent
    For k = 0 To UBound(pin_info.Range_List) - 1
        If d_FixedCurrent >= pin_info.Range_List(k) And d_FixedCurrent <= pin_info.Range_List(k + 1) Then
            pin_info.Init_step = k + 1
            SattleTime = pin_info.WaitTime_List(k)
            If SattleTime > WaitTime Then WaitTime = SattleTime
            Exit For
        End If
    Next k

    thehdw.DCVS.Pins(powerPin).SetCurrentRanges d_FixedCurrent, d_FixedCurrent    ' HexVS
    thehdw.DCVS.Pins(powerPin).CurrentLimit.Source.FoldLimit.level.value = d_FixedCurrent
    
    TheExec.Datalog.WriteComment "All Sites, " & powerPin & ",Fixed Irange: " & d_FixedCurrent
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "FixedIRange_Program_DCVS") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231016][T-BraC][CCT] modify to check IDS efuse name in EFUSE_BitDef_Table sheet
' [20231108][All][Tank] Add Check efuse sheet boolean
Public Sub ParseIDSMappingTable(Optional d_Check_FUSE_Sheet As Boolean = False, Optional ConsiderInstName As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funName As String: funName = "ParseIDSMappingTable"
    
    'string
    Dim sheetName As String
    Dim ParsingArr() As Variant
    Dim s_ErrorMsg As String
    'long
    Dim i As Long
    Dim SheetMaxRow As Long
    Dim SheetMaxCol As Long
    Dim RowLoop As Long
    Dim StageIndex As Long
    Dim m_start As Long
    Dim m_end As Long
    Dim m_loop As Long
    Dim l_sheetCnt As Long
    Dim dic_EfuseIDSName As New Dictionary
    Dim key As Variant
    Dim lEfuseField As eFuseBdfField
    
    Dim b_Sheet_Existing As Boolean
    Const ColStage = 1
    Const ColPinName = 2
    Const ColFieldName = 3
    Const instancename = 4 'Added for consider InstName in mapping table 20240612
    
    sheetName = vbNullString
    s_ErrorMsg = vbNullString
    
    If Flag_IDSMappingTable = False Then
        ''-------------------------
        ''Parse PreCondition Table
        ''-------------------------
        b_Sheet_Existing = False
        sheetName = "IDS_Mapping_Table"
        
        l_sheetCnt = ActiveWorkbook.Sheets.Count
        For i = 1 To l_sheetCnt
            If LCase(Sheets(i).name) Like LCase(sheetName) Then
                b_Sheet_Existing = True
                Exit For
            Else
            End If
        Next i
        
        ''----------
        ''Find stage
        ''----------
        If b_Sheet_Existing Then
            Call GetSheetInfo(sheetName, SheetMaxRow, SheetMaxCol, ParsingArr)
            
            RowLoop = 2
            StageIndex = 0
            Do
                If ParsingArr(RowLoop, ColStage) <> "" Then
                    ReDim Preserve IDS_MAPPING(StageIndex)
                    IDS_MAPPING(StageIndex).stage = ParsingArr(RowLoop, ColStage)
                    IDS_MAPPING(StageIndex).StageCell = RowLoop
                    If ConsiderInstName = True Then
                        IDS_MAPPING(StageIndex).instancename = ParsingArr(RowLoop, instancename) 'Added for consider InstName in mapping table 20240612
                    Else
                    End If
                    StageIndex = StageIndex + 1
                Else
                End If
                RowLoop = RowLoop + 1
            Loop Until (RowLoop = SheetMaxRow + 1)
            
            ''--------------------------------------
            ''Store pin name and field name to array
            ''--------------------------------------
            For StageIndex = 0 To UBound(IDS_MAPPING)
                If StageIndex = UBound(IDS_MAPPING) Then
                    m_end = SheetMaxRow + 1
                Else
                    m_end = IDS_MAPPING(StageIndex + 1).StageCell
                End If
                
                With IDS_MAPPING(StageIndex)
                    Set .MappingDict = New Dictionary
                    .MappingDict.compareMode = TextCompare
                    m_start = .StageCell
                    .cnt = m_end - m_start
                    
                    m_loop = 0          'loop store array
                    RowLoop = m_start   'loop sheet array
                    While (m_loop + m_start <> m_end)
                        .MappingDict.Add ParsingArr(RowLoop, ColPinName), ParsingArr(RowLoop, ColFieldName)
                        RowLoop = RowLoop + 1
                        m_loop = m_loop + 1
                    Wend
                End With
            Next StageIndex
            Flag_IDSMappingTable = True
            
        Else
            TheExec.Datalog.WriteComment "Could not find the sheet, " & sheetName & ", please check it."
        End If
        
        glb_CheckIDSMappingTable_With_Fuse = False
    Else
    End If
    
    If Flag_IDSMappingTable = True And d_Check_FUSE_Sheet = True And glb_CheckIDSMappingTable_With_Fuse = False Then
        '==== Get EFUSE_BitDef_Table sheet ids pin name to check ====
        For Each key In BdfDataBase.Bank_Cfg.DicAllIdsList.Keys
            Set lEfuseField = key
            dic_EfuseIDSName.Add LCase(lEfuseField.name), 0
        Next
        '==== Get EFUSE_BitDef_Table sheet ids pin name to check ====
        
        '==== Check IDS mapping table fuse name exist in EFUSE_BitDef_Table or not====
        For i = 0 To UBound(IDS_MAPPING)
            For Each key In IDS_MAPPING(i).MappingDict.Keys
                If dic_EfuseIDSName.Exists(LCase(IDS_MAPPING(i).MappingDict(CStr(key)))) = False Then    'If BdfDataBase.Bank_Cfg.DicIds.Exists(LCase(ParsingArr(RowLoop, ColFieldName))) = False Then
                    TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
                    s_ErrorMsg = "Parsing IDS_Mapping_Table data not the same with Efuse sheet!!"
                    Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_DC", "ParseIDSMappingTable", s_ErrorMsg)
                    Exit Sub
                Else
                End If
            Next key
        Next i
        '==== Check IDS mapping table fuse name exist in EFUSE_BitDef_Table or not====
        glb_CheckIDSMappingTable_With_Fuse = True
    Else
    End If
Exit Sub
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "ParseIDSMappingTable") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Public Function IDS_Store2Dic_Mapping(Core_Power_Pin As String, All_Power_data As PinListData, patt As Pattern, Optional ConsiderInstName As Boolean = False, Optional OverWriteDic As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    
    Dim i As Integer
    Dim j As Integer
    'long
    Dim mLoop As Long
    Dim Pin_Cnt As Long
    Dim MappingIndex As Long
    'string
    Dim Val_Hi() As String
    Dim Val_Lo() As String
    Dim m_dlogstr As String
    Dim PinAry() As String
    Dim Core_PinAry() As String
    Dim Core_PinAry_exist() As String
    Dim Fuse_StoreName_Ary() As String
    'double
    Dim Core_MeasuredDataAry() As New SiteDouble
    Dim CurrentJob_IDSInfo As IDS_Mapping_Info
    'other variable
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim MappingPin As Variant
    gl_IDS_INFO_Dic.RemoveAll
    
    m_dlogstr = vbNullString
    ''------------------
    ''Get flow use-limit
    ''------------------
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    ' if no Use-Limits on this test, FlowLimitsInfo is nothing
    If FlowLimitsInfo Is Nothing Then
        If isDebugMode Then
            TheExec.AddOutput "Could not get the limits info", vbRed, True
        Else
        End If
        Exit Function
    Else
    End If
    
    FlowLimitsInfo.GetLowLimits Val_Lo
    FlowLimitsInfo.GetHighLimits Val_Hi
    
    ''----------------------------------
    ''Split CorePwrPin and FuseFieldName
    ''----------------------------------
    TheExec.DataManager.DecomposePinList Core_Power_Pin, Core_PinAry, Pin_Cnt
    
    ''--------------
    ''Remove N/C pin
    ''--------------
    j = 0
    For i = 0 To Pin_Cnt - 1
        If TheExec.DataManager.ChannelType(Core_PinAry(i)) <> "N/C" Then
            ReDim Preserve Core_PinAry_exist(j)
            ReDim Preserve Core_MeasuredDataAry(j)
            Core_PinAry_exist(j) = Core_PinAry(i)
            Core_MeasuredDataAry(j) = All_Power_data.Pins(UCase(Core_PinAry_exist(j)))
            j = j + 1
        Else
        End If
    Next i

    ''--------------------------------------------------------------
    ''The fused field get the corresponding measured data and limit.
    ''--------------------------------------------------------------
'    TheExec.Datalog.WriteComment "-----Start to save IDS value-----"
    For i = 0 To UBound(IDS_MAPPING)
        If ConsiderInstName = True Then
            If UCase(currentJobName) = UCase(IDS_MAPPING(i).stage) And glb_TestInstance = UCase(IDS_MAPPING(i).instancename) Then 'Add for consider Inst Name 20240612
                CurrentJob_IDSInfo = IDS_MAPPING(i)
                Exit For
            Else
            End If
        Else
            If UCase(currentJobName) = UCase(IDS_MAPPING(i).stage) Then
                CurrentJob_IDSInfo = IDS_MAPPING(i)
                Exit For
            Else
            End If
        End If
    Next i
    Dim l_MappingDict As Long
    l_MappingDict = CurrentJob_IDSInfo.MappingDict.Count - 1
    ReDim ids_info_ary(l_MappingDict)
    
    
    
    mLoop = 0

    Dim v_PinTemp As Variant
    Dim s_MergePin As String
    Dim pl_MergePin_Temp As New SiteDouble
    For Each v_PinTemp In CurrentJob_IDSInfo.MappingDict.Keys
        s_MergePin = vbNullString
        pl_MergePin_Temp = 0
        TheExec.DataManager.DecomposePinList CStr(v_PinTemp), PinAry, Pin_Cnt
        If Pin_Cnt > 1 Then ''VDD_Fixed = VDD_Fixed_1+VDD_Fixed_2
            For i = 0 To Pin_Cnt - 1
                For j = 0 To UBound(Core_PinAry_exist)
                    If LCase(PinAry(i)) = LCase(Core_PinAry_exist(j)) Then
                        If s_MergePin = "" Then
                            s_MergePin = Core_PinAry_exist(j)
                        Else
                            s_MergePin = s_MergePin & "," & Core_PinAry_exist(j)
                        End If
                        pl_MergePin_Temp = pl_MergePin_Temp.Add(All_Power_data.Pins(UCase(Core_PinAry_exist(j))).value)
                        Exit For
                    Else
                    End If
                Next j
            Next i
        Else
            For j = 0 To UBound(Core_PinAry_exist)
                If LCase(v_PinTemp) = LCase(Core_PinAry_exist(j)) Then
                    s_MergePin = Core_PinAry_exist(j)
                    pl_MergePin_Temp = pl_MergePin_Temp.Add(All_Power_data.Pins(UCase(Core_PinAry_exist(j))))
                    Exit For
                Else
                End If
            Next j
        End If
        
        With ids_info_ary(mLoop)
            .Pat = patt
            .hiLimit = Val_Hi((mLoop)) 'fixed index from j to mloop 20240612
            .loLimit = Val_Lo((mLoop)) 'fixed index from j to mloop 20240612
            .pin = s_MergePin
            .MeasureValue = pl_MergePin_Temp
        End With
        ''Dictionary stores value.
        ''Example ==> gl_IDS_INFO_Dic.Add "ids_vdd_ave", Array(Pin, PAT, MeasureValue, LoLimit, HiLimit)
        ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(0) => "VDD_AVE"
        ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(1) => "pp_scya0_s_fulp_io_popx_nan_daa_dio_allfv_si_ids"
        ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(2) => All_Power_data_IDS_GB.Pins(i)
        ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(3) => "0"
        ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(4) => "271.557"
        If Not gl_IDS_INFO_Dic.Exists(CStr(v_PinTemp)) Then
            gl_IDS_INFO_Dic.Add CStr(v_PinTemp), _
                 Array(ids_info_ary(mLoop).pin, ids_info_ary(mLoop).Pat, ids_info_ary(mLoop).MeasureValue, ids_info_ary(mLoop).loLimit, ids_info_ary(mLoop).hiLimit)
        Else
            If OverWriteDic = True Then
                gl_IDS_INFO_Dic.Remove CStr(v_PinTemp)
                gl_IDS_INFO_Dic.Add CStr(v_PinTemp), _
                    Array(ids_info_ary(mLoop).pin, ids_info_ary(mLoop).Pat, ids_info_ary(mLoop).MeasureValue, ids_info_ary(mLoop).loLimit, ids_info_ary(mLoop).hiLimit)
            Else
                TheExec.Datalog.WriteComment "The fused name of argument is conflicted"
            End If
        End If

        mLoop = mLoop + 1
    Next v_PinTemp
'    TheExec.Datalog.WriteComment "-----End to save IDS value-----"
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "IDS_Store2Dic_Mapping") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230908][T-BraC][CC] Use AddStoredData to store SiteDouble
Public Function Calc_IDS_Sum(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim IDS_String As String
    Dim IDS_Array() As String
    Dim ids_name As Variant
    Dim Dict_Name As String
    Dim Tname_LimitIndex As String
    
    Dim IDSTname As String
    
    Dim Sum_Result As New SiteDouble
    
    Dim tmpArr() As String
    Dim tmpArr2() As String
    Dim TempArr3() As String
    Dim SumType As String
    Dim PinName As String
    Dim SumItem As String
    Dim ItemArr() As String
    Dim sTempBinCutNum As String
    Dim nBinCutNumber As Long
    Dim ItemName As Variant
    Dim opbank As New eFuseBdfBank
    Dim field As New eFuseBdfField
    'Add for group pin name------------
    Dim i As Long
    Dim dummy As Long: dummy = 0
    Dim Fuse_StoreName As String
    Dim aaa As New SiteDouble
   ' Dim CurrentJob_IDSInfo As IDS_Mapping_Info
    Dim Hilimit_IDS As Double
    Dim Lolimit_IDS As Double
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    Dim Val_Hi() As String
    Dim Val_Lo() As String
    Dim site As Variant 'Carter, 20240304
    FlowLimitsInfo.GetHighLimits Val_Hi
    FlowLimitsInfo.GetLowLimits Val_Lo
    'Add for group pin name------------
    tmpArr = Split(argv(0), "@")
    SumType = tmpArr(0)
    If InStr(tmpArr(1), "#") <> 0 Then
        tmpArr2 = Split(tmpArr(1), "#")
        PinName = tmpArr2(0)
        SumItem = tmpArr2(1)
    Else
        SumItem = tmpArr(1)
    End If
    sTempBinCutNum = "999"     'default max
    If InStr(SumItem, "~") Then         'to get specific production identifier
        TempArr3 = Split(SumItem, "~")
        SumItem = TempArr3(0)
        sTempBinCutNum = TempArr3(1)
    End If
    
    ItemArr = Split(SumItem, "+")
    
    If UBound(argv) > 0 Then Dict_Name = argv(1)
    
    Select Case SumType
        Case "EFUSE"
            Set opbank = GetBdfBank("CFG")
            For Each ItemName In ItemArr
                Set field = opbank.Fields(ItemName)
                Sum_Result = Sum_Result.Add(field.DsscDecValue.Multiply(field.Resolution * 0.001))
                TheExec.Flow.TestLimit field.DsscDecValue.Multiply(field.Resolution * 0.001), ForceResults:=tlForceNone, Tname:=ItemName, PinName:=ItemName
            Next ItemName
            If TheExec.TesterMode = testModeOffline Then Sum_Result = 0.01 + Rnd() * 0.0001
            Call AddStoredData(PinName + "_EFUSE", Sum_Result)
            TheExec.Flow.TestLimit Sum_Result, ForceResults:=tlForceNone, Tname:="Sum_Result"
        Case "IDS"
            For Each ItemName In ItemArr
                Sum_Result = Sum_Result.Add(All_Power_data_IDS_GB.Pins(ItemName))
            Next ItemName
            Call GetFlowTName
            
            Call AddStoredData(PinName, Sum_Result)
           
            'Add for group pin name------------
            IDSTname = Report_TName_From_Instance(CalcI, "X")
            
            Dim nTempBinCutNum As New SiteLong
    
            For Each site In TheExec.sites
            
                If sTempBinCutNum = "d" Then        'if "d" mean get device value
                    '//////////////////////////////////////////if is seudofuse//////////////////////////////////////////
                    Dim field_normal As eFuseBdfField
        
                    If PseudoFuseEnable Then
                        If BdfDataBase.Banks("CFG").Fields.Exists("product_identifier") Then
                            Set field_normal = BdfDataBase.Banks("CFG").Fields("product_identifier")
                        
                            CurrentPassBinCutNum_normal(site) = field_normal.DsscDecValue + 1
                            If CurrentPassBinCutNum_normal(site) > Total_Bincut_Num Then
                            GlbUtility.WriteDlg "site:" & site & ", product_identifier " & CurrentPassBinCutNum_normal(site) & " > Total_Bincut_Num " & Total_Bincut_Num & " , Error!!!"
                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="product_identifier Error"
                            End If
                        Else
                            CurrentPassBinCutNum_normal = 1
                        End If
                    End If
                    '//////////////////////////////////////////if is seudofuse//////////////////////////////////////////

                    nBinCutNumber = CurrentPassBinCutNum_normal
                    nTempBinCutNum(site) = CheckHaveBinCutSheet(nBinCutNumber, True)      'check has production identifier reference BinCut sheet
                Else
                    nBinCutNumber = CLng(sTempBinCutNum)
                    nTempBinCutNum(site) = CheckHaveBinCutSheet(nBinCutNumber)      'check has production identifier reference BinCut sheet
                End If
                     
                Hilimit_IDS = Compare_ForceVal_BV(PinName, Val_Hi(TheExec.Flow.TestLimitIndex), True, nBincutNum:=nTempBinCutNum(site))        'if pins without limit use bincut spec IDSmax as current range
                Lolimit_IDS = Compare_ForceVal_BV(PinName, Val_Lo(TheExec.Flow.TestLimitIndex), False, nBincutNum:=nTempBinCutNum(site))
                Lolimit_IDS = IIf(Hilimit_IDS = -999, Hilimit_IDS, Lolimit_IDS)
                
                If Hilimit_IDS = 0 And Lolimit_IDS = 0 Then
                    TheExec.Flow.TestLimit Sum_Result, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=IDSTname, highCompareSign:=tlSignNone, lowCompareSign:=tlSignNone
                ElseIf Hilimit_IDS = 0 Then
                    TheExec.Flow.TestLimit Sum_Result, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=IDSTname, highCompareSign:=tlSignNone, lowVal:=Lolimit_IDS
                ElseIf Lolimit_IDS = 0 Then
                    TheExec.Flow.TestLimit Sum_Result, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=IDSTname, hiVal:=Hilimit_IDS, lowCompareSign:=tlSignNone
                Else
                    TheExec.Flow.TestLimit Sum_Result, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=IDSTname, hiVal:=Hilimit_IDS, lowVal:=Lolimit_IDS
                End If
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
    End Select

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "Calc_IDS_Sum") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Disconnect_X_and_Nouse_Pin(eDisconnectPinType As DisconnectPinType, patt As Pattern, Optional EnableDisconnectPins As Boolean = False, Optional ExcludeDisconnectPins As String = vbNullString, _
                                        Optional sDontCarePin As String = vbNullString, Optional sOutPut_DisconnectPin As String = vbNullString, Optional sOutput_IgnorePin As String = vbNullString, _
                                        Optional sOutput_NCPin As String = vbNullString) As Boolean
On Error GoTo errHandler
    Dim sFuncName As String:: sFuncName = "Disconnect_X_and_Nouse_Pin"
    '====20221121 Tank add Disconnect pin====
    Dim nExcludePinCnt As Long
    Dim excludePinAry() As String
    Dim dicSearchPin As New Dictionary
    'Dim DisconnectPinAry() As String
    Dim sDisconnectPin As String
    Dim nTempCnt As Long
    Dim sIgnorePin As String
    Dim dicSamePin As New Dictionary
    Dim vPin As Variant
    Dim sPinName As String
    Dim n As Long
    Dim s_ErrorStatement As String
    Dim nTestNumber As Long
    Dim sUnUsedPin As String
    Dim nPatternIndex As Long

    If EnableDisconnectPins Then
        If eDisconnectPinType = DisconnectPinType.UnusedIoPin Then    'first disconnect Unused pin before run pattern
        
            Call ParsePreCondition      '20221121 no need eng_IDS_swep
    
            Call Parse_PatSetAll        '20221121 load PatSetAll
            
            Call Check_PreCondition_Pattern_Version     '20221121 compare pattern version
        
            sUnUsedPin = vbNullString
            dicSearchPin.RemoveAll

            ENG_SweepPin = False    'not sweep pin when EnableDisconnectPins

            Select Case gl_ePreConditionError
            
                Case PreConditionError.PatRepetition:
                    s_ErrorStatement = "PreCondition sheet pattern repetition!!"
                    GoTo ProcessBinOut
                Case PreConditionError.PatVerDifferent:
                    s_ErrorStatement = "PreCondition sheet pattern version not match!!"
                    GoTo ProcessBinOut
                Case PreConditionError.SheetNotExist:
                    s_ErrorStatement = "PreCondition sheet not exist!!"
                    GoTo ProcessBinOut
                Case PreConditionError.ParseSheetError:
                    s_ErrorStatement = "PreCondition or PatSetAll sheet parse error!!"
                    GoTo ProcessBinOut
                    
            End Select
            
            '====Check include "exclude disconnect pin" exist====
            TheExec.DataManager.DecomposePinList ExcludeDisconnectPins, excludePinAry(), nExcludePinCnt

            If nExcludePinCnt > 0 Then        'make excludePin that search fast
                For Each vPin In excludePinAry
                    sPinName = vPin
                    dicSearchPin.Add UCase(sPinName), sPinName
                Next vPin
            End If
            '====Check include "exclude disconnect pin" exist====
            
            If gl_PreCondition_Dic.Exists(LCase(patt.value)) Then        'find IDS pattern exist in precondition
                nPatternIndex = gl_PreCondition_Dic.item(LCase(patt.value))
  
                For Each vPin In PreConditionInfo(nPatternIndex).Arr_UnusedPins
                    sPinName = vPin
                    If dicSearchPin.Exists(UCase(sPinName)) = False Then      'pin not same in EnableDisconnectPins need to disconnect
                        If UCase(TheExec.DataManager.ChannelType(sPinName)) <> "N/C" Then      'check pin channel type not N/C
    
                            sDisconnectPin = CombineStringList(sDisconnectPin, UCase(sPinName))       'Combine disconnect pin list string
                            sUnUsedPin = CombineStringList(sUnUsedPin, UCase(sPinName))
                            nTempCnt = nTempCnt + 1
                        Else
                            sOutput_NCPin = CombineStringList(sOutput_NCPin, UCase(sPinName))
                            
                        End If
                    Else
                        dicSamePin.Add UCase(sPinName), sPinName
                    End If
                Next vPin
                
                For Each vPin In PreConditionInfo(nPatternIndex).Arr_XPins
                    sPinName = vPin

                    If dicSearchPin.Exists(UCase(sPinName)) = False Then      'pin not same in EnableDisconnectPins need to disconnect
                        If UCase(TheExec.DataManager.ChannelType(sPinName)) <> "N/C" Then      'check pin channel type not N/C
    
                            sDisconnectPin = CombineStringList(sDisconnectPin, UCase(sPinName))       'Combine disconnect pin list string
                            sDontCarePin = CombineStringList(sDontCarePin, UCase(sPinName))
                            nTempCnt = nTempCnt + 1
                        Else
                            sOutput_NCPin = CombineStringList(sOutput_NCPin, UCase(sPinName))
                        End If
                    Else
                        dicSamePin.Add UCase(sPinName), sPinName
                    End If
                Next vPin

            Else
                s_ErrorStatement = "IDS test pattern not found in PreCondition sheet!!"
                GoTo ProcessBinOut
            End If
            
            If sUnUsedPin <> "" Then thehdw.Digital.Pins(sUnUsedPin).Disconnect      'Disconnect UnUsed IO pin parallel
            
            TheExec.Datalog.WriteComment "<Warning>: " & "Disconnect UnUse pin and Xpin : " & sDisconnectPin
            sOutPut_DisconnectPin = sDisconnectPin
            
            If nExcludePinCnt > 0 Then
                For Each vPin In excludePinAry      'find ExcludeDisconnectPins not in PreConDiction sheet
                    If dicSamePin.Exists(UCase(vPin)) = False Then
                        sIgnorePin = CombineStringList(sIgnorePin, UCase(vPin))
                    End If
                Next vPin
                
                If sIgnorePin <> "" Then
                    sOutput_IgnorePin = sIgnorePin
                    s_ErrorStatement = "include ExcludeDisconnectPins argument not in PreCondition sheet pins : " & sIgnorePin
                    GoTo ProcessBinOut
                End If

            End If
        ElseIf eDisconnectPinType = DisconnectPinType.Xpin Then   'second disconnect Xpin after run pattern
        
            If sDontCarePin <> "" Then thehdw.Digital.Pins(sDontCarePin).Disconnect      'Disconnect Don't care pin
            
        End If
        
        '====20221121 Tank add Disconnect pin====
    End If
    Disconnect_X_and_Nouse_Pin = True
    Exit Function
ProcessBinOut:  'Binout if get something error
    TheExec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="Disconnect_X_and_Nouse_Pin Error"
    Call Print_Error_Message(Warning_Info, LIB_DC, sFuncName, s_ErrorStatement)
    Disconnect_X_and_Nouse_Pin = False
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_DC, sFuncName)
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20230407][T-All] testlimit compare power short result
Public Function DCVS_PowerPinContiCompare(pininfo As PinData, sPinName As String, TmpTname As String, ForceV As Double, TestLimitMode As tlLimitForceResults, LowLimit As Double, hiLimit As Double)
On Error GoTo errHandler
    Dim sFuncName As String:: sFuncName = "DCVS_PowerPinContiCompare"
    Dim site As Variant 'Carter, 20240304
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            pininfo.value(site) = 0.01 + Rnd() * 0.0001
        Next site
    End If
    
    If TestLimitMode = tlForceFlow Then
        TheExec.Flow.TestLimit resultVal:=pininfo, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TmpTname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
    ElseIf TestLimitMode = tlForceNone Then
        TheExec.Flow.TestLimit resultVal:=pininfo, lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TmpTname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
    End If
    TheExec.Datalog.WriteComment "Current I Range: " & sPinName & "--->" & thehdw.DCVS.Pins(sPinName).Meter.CurrentRange.value
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_DC, sFuncName)
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][CP] To give true flag name to set
' [20231228][T-All][Tank] if pin NC no need to compare limit
Public Function CheckTestInst_HiLoLimit(LowLimit As Double, hiLimit As Double, Lolimit_new() As Double, HiLimit_new() As Double, Pins() As String, _
                                        TestLimitMode As tlLimitForceResults, Optional Flag_Open As String = "F_open", Optional FlowLimitIndex As Integer = 0) As Boolean

    On Error GoTo errHandler
    Dim sFuncName As String:: sFuncName = "CheckTestInst_HiLoLimit"
    Dim i As Long
    Dim DUTPin As Variant
    Dim site As Variant 'Carter, 20240304
    
    CheckTestInst_HiLoLimit = False
    '---------------------------------------------------------'
    '--- Lowlimit value should be more than ClampLow value ---'
    '--- Highlimit value should be less than ClampHigh value -'
    '---------------------------------------------------------'
    If TestLimitMode = tlForceFlow Then
        i = 0 + FlowLimitIndex
        For Each DUTPin In Pins
            If gl_GetInstrument_Dic.Exists(LCase(DUTPin)) Then
                If Lolimit_new(i) <= thehdw.PPMU.Pins(DUTPin).ClampVLo.value Then
                    Call Print_Error_Message(Error_Info, LIB_DC, sFuncName, TheExec.DataManager.instancename & ", Pin:" & DUTPin & " LowLimit(" & Lolimit_new(i) & "V) <= ClampVLo(" & thehdw.PPMU.Pins(DUTPin).ClampVLo.value & "V)")
                    For Each site In TheExec.sites
                        TheExec.sites.item(site).FlagState(Flag_Open) = logicTrue
                    Next
                    gl_isCheckClampLimit = ContiClampCheckType.CheckFail
                    CheckTestInst_HiLoLimit = True
                    Exit Function
                ElseIf HiLimit_new(i) >= thehdw.PPMU.Pins(DUTPin).ClampVHi.value Then
                    Call Print_Error_Message(Error_Info, LIB_DC, sFuncName, TheExec.DataManager.instancename & ", Pin:" & DUTPin & " HiLimit(" & HiLimit_new(i) & "V) >= ClampVHi(" & thehdw.PPMU.Pins(DUTPin).ClampVHi.value & "V)")
                    For Each site In TheExec.sites
                        TheExec.sites.item(site).FlagState(Flag_Open) = logicTrue
                    Next
                    gl_isCheckClampLimit = ContiClampCheckType.CheckFail
                    CheckTestInst_HiLoLimit = True
                    Exit Function
                Else
                End If
            End If
            i = i + 1
        Next DUTPin
        FlowLimitIndex = i
        
    ElseIf TestLimitMode = tlForceNone Then
        For Each DUTPin In Pins
            If LowLimit <= thehdw.PPMU.Pins(DUTPin).ClampVLo.value Then
                Call Print_Error_Message(Error_Info, LIB_DC, sFuncName, TheExec.DataManager.instancename & ", Pin:" & DUTPin & " LowLimit(" & LowLimit & "V) <= ClampVLo(" & thehdw.PPMU.Pins(DUTPin).ClampVLo.value & "V)")
                For Each site In TheExec.sites
                    TheExec.sites.item(site).FlagState(Flag_Open) = logicTrue
                Next
                gl_isCheckClampLimit = ContiClampCheckType.CheckFail
                CheckTestInst_HiLoLimit = True
                Exit Function
            ElseIf hiLimit >= thehdw.PPMU.Pins(DUTPin).ClampVHi.value Then
                Call Print_Error_Message(Error_Info, LIB_DC, sFuncName, TheExec.DataManager.instancename & ", Pin:" & DUTPin & " HiLimit(" & hiLimit & "V) >= ClampVHi(" & thehdw.PPMU.Pins(DUTPin).ClampVHi.value & "V)")
                For Each site In TheExec.sites
                    TheExec.sites.item(site).FlagState(Flag_Open) = logicTrue
                Next
                gl_isCheckClampLimit = ContiClampCheckType.CheckFail
                CheckTestInst_HiLoLimit = True
                Exit Function
            Else
            End If
        Next DUTPin
    Else
    End If
                                                                                                                                                                                                                         
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_DC, sFuncName)
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function Measure_Pin_FIMV(s_MeasurePin As String, d_forceI As Double, Optional b_isPN_Disconnect As Boolean = False, Optional d_WaitTime As Double = 0.005) As PinListData
On Error GoTo errHandler
    If UCase(gl_GetInstrumentType_Dic(LCase(s_MeasurePin))) Like "*DCVS*" Then
        Set Measure_Pin_FIMV = DCVS_Pin_FIMV(s_MeasurePin, d_forceI, d_WaitTime)
    ElseIf UCase(gl_GetInstrumentType_Dic(LCase(s_MeasurePin))) Like "*DCVI*" Then
        Set Measure_Pin_FIMV = DCVI_Pin_FIMV(s_MeasurePin, d_forceI, d_WaitTime)
    ElseIf UCase(gl_GetInstrumentType_Dic(LCase(s_MeasurePin))) = "I/O" Then
        Set Measure_Pin_FIMV = Digital_Pin_FIMV(s_MeasurePin, d_forceI, b_isPN_Disconnect, d_WaitTime)
    End If
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_DC", "Measure_Pin_FIMV")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function Digital_Pin_FIMV(s_MeasurePin As String, d_forceI As Double, b_isPN_Disconnect As Boolean, Optional d_WaitTime As Double = 0.005) As PinListData
On Error GoTo errHandler
    Dim s_PinStr As String
    Dim tmpPinListData As New PinListData
    
    If UCase(gl_GetInstrument_Dic(LCase(s_MeasurePin))) = glbConstIns_UP2200 Or UCase(gl_GetInstrument_Dic(LCase(s_MeasurePin))) = glbConstIns_UP1600 Then
        tmpPinListData.AddPin (s_MeasurePin)
        With thehdw.PPMU.Pins(s_MeasurePin)
            .Gate = tlOff
            .ForceI (d_forceI)
            .Gate = tlOn
        End With

        thehdw.Wait d_WaitTime
    
        If b_isPN_Disconnect = False Then
            DebugPrintFunc_PPMU CStr(s_MeasurePin)
            tmpPinListData.Pins(s_MeasurePin) = thehdw.PPMU.Pins(s_MeasurePin).Read(tlPPMUReadMeasurements, 20)    'normal measure
        Else
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////
            If DicDiffPairs.Exists(LCase(CStr(s_MeasurePin))) Then
                 s_PinStr = DicDiffPairs(LCase(CStr(s_MeasurePin)))
                 thehdw.PPMU.Pins(s_PinStr).Gate = tlOff
                 DebugPrintFunc_PPMU CStr(s_MeasurePin)
                 tmpPinListData.Pins(s_MeasurePin) = thehdw.PPMU.Pins(s_MeasurePin).Read(tlPPMUReadMeasurements, 32)
                 thehdw.PPMU.Pins(s_PinStr).Gate = tlOn   'recover
            Else
                  DebugPrintFunc_PPMU CStr(s_MeasurePin)
                  tmpPinListData.Pins(s_MeasurePin) = thehdw.PPMU.Pins(s_MeasurePin).Read(tlPPMUReadMeasurements, 32)
            End If

        End If

        'recover measure dut Pin to 0V before next Pin
        thehdw.PPMU.Pins(s_MeasurePin).ForceV (0) 'correct it to force v, not force i.
        Set Digital_Pin_FIMV = tmpPinListData
    End If
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_DC", "Digital_Pin_FIMV")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function DCVI_Pin_FIMV(s_MeasurePin As String, d_forceI As Double, Optional d_WaitTime As Double = 0.005) As PinListData
On Error GoTo errHandler
    Dim tmpPinListData As New PinListData
    
    'Added for deal special pins 20240702 michael
    Dim Special_pin_pair As String

    If glb_SpecialPin_dic.Exists(s_MeasurePin) = True Then
        Special_pin_pair = glb_SpecialPin_dic.item(s_MeasurePin)
        With thehdw.DCVI.Pins(Special_pin_pair)
            .Gate = False
            .Disconnect
        End With
    End If
    
    tmpPinListData.AddPin (s_MeasurePin)
    With thehdw.DCVI.Pins(s_MeasurePin)
        .Disconnect tlDCVIConnectDefault
        .mode = tlDCVIModeCurrent
        .Voltage = 1    'original setting is clamp voltage=0 which is not correct 2018/01/02
        .VoltageRange.value = -2
        .Current = d_forceI
        .CurrentRange.Autorange = True
        .Connect tlDCVIConnectDefault
        .Gate = True
        
        'measure
        .Meter.mode = tlDCVIMeterVoltage
    End With
    
    thehdw.Wait d_WaitTime ' TTR from 0.5 to 0.03 'to 5ms 180430
    
    If UCase(gl_GetInstrument_Dic(LCase(s_MeasurePin))) = glbConstIns_DC30 Then
        tmpPinListData.Pins(s_MeasurePin) = thehdw.DCVI.Pins(s_MeasurePin).Meter.Read(tlStrobe, 1) '''only for DC30
    Else
        tmpPinListData.Pins(s_MeasurePin) = thehdw.DCVI.Pins(s_MeasurePin).Meter.Read(tlStrobe, 1, 1000000)
    End If
    
    'reset
    With thehdw.DCVI.Pins(s_MeasurePin)
        .mode = tlDCVIModeVoltage
        .Voltage = 0
    End With

    If glb_SpecialPin_dic.Exists(s_MeasurePin) = True Then
        Special_pin_pair = glb_SpecialPin_dic.item(s_MeasurePin)
        With thehdw.DCVI.Pins(Special_pin_pair)
            .mode = tlDCVIModeVoltage
            .Voltage = 0
            .Connect tlDCVIConnectDefault
            .Gate = True
        End With
    End If
    
    Set DCVI_Pin_FIMV = tmpPinListData
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_DC", "DCVI_Pin_FIMV")
    If AbortTest Then Exit Function Else Resume Next
End Function

'added to store special pins pair 20240702 michael
Public Function Store_Special_pins_pair(pin As String)
On Error GoTo errHandler
    
    If pin Like "*TD_B" And glb_SpecialPin_dic.Exists(pin) = False Then
        glb_SpecialPin_dic.Add pin, pin & "S"
    ElseIf pin Like "*TD_BS" And glb_SpecialPin_dic.Exists(pin) = False Then
        glb_SpecialPin_dic.Add pin, left(pin, Len(pin) - 1)
    ElseIf pin Like "*TD_E" And glb_SpecialPin_dic.Exists(pin) = False Then
        glb_SpecialPin_dic.Add pin, pin & "S"
    ElseIf pin Like "*TD_ES" And glb_SpecialPin_dic.Exists(pin) = False Then
        glb_SpecialPin_dic.Add pin, left(pin, Len(pin) - 1)
    End If
    
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_DC", "Store_Special_pins_pair")
    If AbortTest Then Exit Function Else Resume Next
End Function
' [20231003][All][Tank] modify after Chihome review
Public Function DCVS_Pin_FIMV(s_MeasurePin As String, d_forceI As Double, Optional d_WaitTime As Double = 0.005) As PinListData
On Error GoTo errHandler
    Dim tmpPinListData As New PinListData
    
    Dim Special_pin_pair As String

    If glb_SpecialPin_dic.Exists(s_MeasurePin) = True Then
        Special_pin_pair = glb_SpecialPin_dic.item(s_MeasurePin)
        With thehdw.DCVS.Pins(Special_pin_pair)
            .Gate = False
            .Disconnect
        End With
    End If
    
    tmpPinListData.AddPin (s_MeasurePin)
    With thehdw.DCVS.Pins(s_MeasurePin)
        .Disconnect
        
        .CurrentRange.value = Abs(d_forceI)
        .CurrentLimit.Source.FoldLimit.level.value = Abs(d_forceI)
        .CurrentLimit.Sink.FoldLimit.level.value = Abs(d_forceI)
               
        If glb_TesterType = "UltraFLEXplus" Then
            .mode = tlDCVSModeHighImpedance
        Else
            .mode = tlDCVSModeVoltage
        End If
        
        If d_forceI > 0 Then
            .Voltage.value = 1.5
        Else
            If glb_TesterType = "UltraFLEXplus" Then
                .VoltageRange.value = 5.5
            End If
            .Voltage.value = -1.5
        End If
   

        .Meter.mode = tlDCVSMeterVoltage

        .Connect
        .Gate = True
        
        If glb_TesterType = "UltraFLEXplus" Then
            .mode = tlDCVSModeCurrent
        End If
    End With

    thehdw.Wait d_WaitTime 'follow Cyprus to do DC TTR by KWLU 180601

    tmpPinListData.Pins(s_MeasurePin) = thehdw.DCVS.Pins(s_MeasurePin).Meter.Read(tlStrobe)
    
    'reset
    If glb_TesterType = "UltraFLEXplus" Then
        thehdw.DCVS.Pins(s_MeasurePin).mode = tlDCVSModeHighImpedance    'tlDCVSModeVoltage
    End If
    thehdw.DCVS.Pins(s_MeasurePin).Voltage.value = 0#
    
    If glb_SpecialPin_dic.Exists(s_MeasurePin) = True Then
        Special_pin_pair = glb_SpecialPin_dic.item(s_MeasurePin)
        With thehdw.DCVS.Pins(Special_pin_pair)
            If glb_TesterType = "UltraFLEXplus" Then
                .mode = tlDCVSModeHighImpedance
            Else
                .mode = tlDCVIModeVoltage
            End If
            .Voltage = 0
            .Connect tlDCVSConnectDefault
            .Gate = True
        End With
    End If
    
    Set DCVS_Pin_FIMV = tmpPinListData
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_DC", "DCVS_Pin_FIMV")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function Calc_ActuralClampCurrent(s_DUTPin As String, d_forceI As Double, d_ActuralForceI As Double, s_ErrorMsg As String) As Boolean
On Error GoTo errHandler
    Select Case (UCase(gl_GetInstrument_Dic(LCase(s_DUTPin))))
        Case glbConstIns_VS5A, glbConstIns_VHDVS, glbConstIns_VS800MA
            '''Actural Force Current = Force_I + CurrentRange * 12.5%
            d_ActuralForceI = d_forceI - (0.125 * thehdw.DCVS.Pins(s_DUTPin).CurrentRange.value)
        Case glbConstIns_HEXVS, glbConstIns_VSM
            d_ActuralForceI = d_forceI
        Case Else
            '// if instrument no define binout
            Calc_ActuralClampCurrent = True
    End Select
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_DC", "Calc_ActuralClampCurrent")
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20240110][All][Tank] Add store argument fixcurrentrange i dictionary
Public Function FixCurrentRange_StrToDic(FixCurrentRange As String, dict_FixedCurrent As Dictionary)

On Error GoTo errHandler
    Dim i As Long
    Dim j As Long
    Dim s_FixedCurrent_Ary() As String
    Dim s_FixedCurrent As String
    Dim s_FixedCurrent_Pin As String
    Dim s_FixedCurrent_PinAry() As String
    Dim Pin_Cnt As Long
   
    s_FixedCurrent = vbNullString
    s_FixedCurrent_Pin = vbNullString
    
    If FixCurrentRange <> "" Then  'FixCurrentRange 20220627
        s_FixedCurrent_Ary = SplitInputCondition(FixCurrentRange, ";")
        For i = 0 To UBound(s_FixedCurrent_Ary)
            If InStr(s_FixedCurrent_Ary(i), ":") <> 0 Then
                s_FixedCurrent = Split(s_FixedCurrent_Ary(i), ":")(0)
                s_FixedCurrent_Pin = Split(s_FixedCurrent_Ary(i), ":")(1)
                's_FixedCurrent_PinAry = SplitInputCondition(s_FixedCurrent_Pin, ",")
                'Add to support Pingroup 20240612
                TheExec.DataManager.DecomposePinList s_FixedCurrent_Pin, s_FixedCurrent_PinAry, Pin_Cnt

                For j = 0 To UBound(s_FixedCurrent_PinAry)
                    If dict_FixedCurrent.Exists(LCase(s_FixedCurrent_PinAry(j))) Then
                        TheExec.Datalog.WriteComment "Pin:" & LCase(s_FixedCurrent_PinAry(j)) & " -> Fixcurrent already in dictionary"
                    Else
                        dict_FixedCurrent.Add LCase(s_FixedCurrent_PinAry(j)), s_FixedCurrent
                    End If
                Next j
            Else
                TheExec.Datalog.WriteComment "There is no argument for Fixed Current Range setting, Please check the instance."
            End If
        Next i
    Else
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_DC, "FixCurrentRange_StrToDic")
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20240110][All][Tank] add setcurretrange subfunction
Public Function SetGroupCurrentRange(PinsGroup As String, Grp_Val_Hi As String, FlowLimitInit As Boolean, WaitTime As Double, TestHiLimit As String, _
                                Optional dict_FixedCurrent As Dictionary)
    
    Dim i As Long
    Dim j As Long
    
    Dim Irange As Double
    Dim SettleTime As Double
    
    Dim PinsGroup_Ary() As String
    Dim Grp_Val_Hi_Ary() As String
    Dim sa_HiLimit_Ary() As String
    
    Dim funcName As String: funcName = "SetCurrentRange"
    Dim s_PinInstrumentTyp As String
    
    On Error GoTo errHandler
    
    PinsGroup_Ary = Split(PinsGroup, ",")
    Grp_Val_Hi_Ary = Split(Grp_Val_Hi, ",")
    sa_HiLimit_Ary = Split(TestHiLimit, ",")
    
    For i = 0 To UBound(PinsGroup_Ary)
        If gl_GetInstrumentType_Dic.Exists(LCase(PinsGroup_Ary(i))) Then
            If dict_FixedCurrent.Exists(LCase(PinsGroup_Ary(i))) Then
                Irange = CDbl(dict_FixedCurrent(LCase(PinsGroup_Ary(i))))
            Else
                If (FlowLimitInit = True) Then
                    Irange = CDbl(Grp_Val_Hi_Ary(i))
                Else
                    Irange = CDbl(sa_HiLimit_Ary(i))
                End If
            End If
            
            Call SetCurrentRange(LCase(PinsGroup_Ary(i)), Irange, WaitTime)
        End If
    Next i
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "SetCurrentRange") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20240110][All][Tank] add setcurretrange subfunction
Public Function SetCurrentRange(powerPin As String, SetCurrentValue As Double, WaitTime As Double, Optional SetCurrentStep As Long)
On Error GoTo errHandler
    Dim k As Long
    Dim SattleTime As Double
    Dim s_InstType As String
    
    If gl_GetInstrumentType_Dic.Exists(LCase(powerPin)) Then
        s_InstType = gl_GetInstrumentType_Dic(LCase(powerPin))
        With PowerPin_range_ary(gl_dicPowerPinIndex(LCase(powerPin)))
            If s_InstType Like "*DCVS*" Then
                For k = 0 To UBound(.Range_List)
                    If (SetCurrentValue <= .Range_List(k)) Or (k = UBound(.Range_List)) Then
                        thehdw.DCVS.Pins(powerPin).SetCurrentRanges .Range_List(k), .Range_List(k)
                        If .Range_List(k) < PowerPin_range_ary(gl_dicPowerPinIndex(LCase(powerPin))).MinIFoldLimit Then
                            thehdw.DCVS.Pins(powerPin).CurrentLimit.Source.FoldLimit.level.value = PowerPin_range_ary(gl_dicPowerPinIndex(LCase(powerPin))).MinIFoldLimit
                        Else
                            thehdw.DCVS.Pins(powerPin).CurrentLimit.Source.FoldLimit.level.value = .Range_List(k)
                        End If
                        SetCurrentStep = k
                        SattleTime = .WaitTime_List(k)
                        If SattleTime > WaitTime Then WaitTime = SattleTime
                        Exit For
                    End If
                Next k
            ElseIf s_InstType Like "*DCVI*" Then
                For k = 0 To UBound(.Range_List)
                    If (SetCurrentValue <= .Range_List(k)) Or (k = UBound(.Range_List)) Then
                        thehdw.DCVI.Pins(powerPin).SetCurrentAndRange .Range_List(k), .Range_List(k)
                        SetCurrentStep = k
                        SattleTime = .WaitTime_List(k)
                        If SattleTime > WaitTime Then WaitTime = SattleTime
                        Exit For
                    End If
                Next k
            End If
            TheExec.Datalog.WriteComment "All Sites, " & powerPin & ",SetCurrentRange: " & .Range_List(k)
        End With
    End If

Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "SetCurrentRange") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
