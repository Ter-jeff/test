Attribute VB_Name = "VBT_LIB_Customized"
Option Explicit


Public Flag_DieControlEnable As Boolean
Public Flag_FlagControlEnable As Boolean

Public ControlTableDie() As Dieinfo
'Public ControlTableFlag() As FlagInfo
Public ControlTableFlag_ExpectedTotalDies As Long
Public ControlTableFlag_CollectedTotalDies As Long
Public ControlTableFlag_maxLen As Integer
Public Type Dieinfo
    XCoord As Long
    YCoord As Long
End Type

Public Type FlagInfo
    flagName As String
    FlagState As LogicState
End Type

Enum LogicState
    localTrue = 1
    localFalse = 0
    localClear = -1
End Enum

Public AlreadyParsed As Boolean

'Public DSSCMappingTableIsRead As Boolean
''Carter - duplicate
''Public Type TestCondition
''    DigSrc_BinStr As String
''    ConditionName As String
''    DigSrc_BitCount As Double
''End Type
''
''Public Type DynamicSrc
''    PatternName As String
''    TestCase() As TestCondition
''End Type
''
''Public SrcStock() As DynamicSrc

''Public Previous_DCCategory As String
''Public Previous_DCSelector As String

Public ControlTableDic As New Dictionary
Dim ControlTableArrayFlag() As Variant


Public Pre_LotTmp As String
Public Pre_WaferIDTmp As String

Public Function Enable_SubFlowFlag_ByDieMap() As Long
'SWLIN 2017/03/29
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Enable_SubFlowFlag_ByDieMap"
    Dim site As Variant
    Dim i As Long
    Dim CompareXcoord As New SiteLong
    Dim CompareYcoord As New SiteLong
    Dim Key_ProberXYcoord As String
    
        
     
        If Flag_DieControlEnable = True Then
            '--- initialize flag and get prober XY coordinate---
            TheExec.Datalog.WriteComment "----------------------------------------"
            For Each site In TheExec.sites
                CompareXcoord(site) = 0: CompareYcoord(site) = 0 ' initlize
                
                If LCase(currentJobName) Like "*cp*" Then
                    CompareXcoord(site) = XCoord(site) 'from prober
                    CompareYcoord(site) = YCoord(site) 'from prober
                ElseIf LCase(currentJobName) Like "*ft*" Then
                    If LCase(currentJobName) Like "*wlft*" And TheExec.enableWord("STDC_ProberXY4WLFT") = True Then
                        'STDC_ProberXY4WLFT, Sampling Test Data Collection is on
                        CompareXcoord(site) = XCoord(site) 'from prober
                        CompareYcoord(site) = YCoord(site) 'from prober
                        TheExec.Datalog.WriteComment ""
                        TheExec.Datalog.WriteComment "------------------------------------------------------------"
                        TheExec.Datalog.WriteComment "Start to compare Die info between 'prober' and 'Control Table'"
                        TheExec.Datalog.WriteComment "------------------------------------------------------------"
                    Else
                        CompareXcoord(site) = HramXCoord(site) 'from efuse
                        CompareYcoord(site) = HramYCoord(site) 'from efuse
                        TheExec.Datalog.WriteComment ""
                        TheExec.Datalog.WriteComment "------------------------------------------------------------"
                        TheExec.Datalog.WriteComment "Start to compare Die info between 'efuse' and 'Control Table'"
                        TheExec.Datalog.WriteComment "------------------------------------------------------------"
                    End If
                End If
                
                Key_ProberXYcoord = CStr(CompareXcoord(site)) + "," + CStr(CompareYcoord(site))
                    
                If ControlTableDic.Exists(Key_ProberXYcoord) Or TheExec.flow.enableWord("Force_9Site_DataCollection") = True Then
                    TheExec.sites(site).FlagState("SamplingTest_TurnOnFlag_ByDie") = logicTrue
                    TheExec.Datalog.WriteComment "Site(" & site & ")" & ", (X,Y) = (" & CompareXcoord(site) & "," & CompareYcoord(site) & "), (Match up)"
                Else
                    TheExec.sites(site).FlagState("SamplingTest_TurnOnFlag_ByDie") = logicFalse
                    TheExec.Datalog.WriteComment "Site(" & site & ")" & ", (X,Y) = (" & CompareXcoord(site) & "," & CompareYcoord(site) & ")"
                End If
                
            Next site
            TheExec.Datalog.WriteComment "----------------------------------------"
        End If
        
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Enable_SubFlowFlag_ByFlagMap() As Long
'SWLIN 2017/03/29
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Enable_SubFlowFlag_ByFlagMap"
    Dim site As Variant
    Dim i As Long
    Dim ComparedFlagName As String
    Dim ComparedFlagState As Variant
    Dim ComparedFlagState_Str As String
    Dim TestResultFlagState_Str As String
    Dim CompareXcoord As New SiteLong
    Dim CompareYcoord As New SiteLong
    
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "-------------------------------------------------------------------"
    TheExec.Datalog.WriteComment "Start to compare Flag-State between 'Tset Result' and 'Control Table'"
    TheExec.Datalog.WriteComment "-------------------------------------------------------------------"
            
    If Flag_FlagControlEnable = True Then
        For Each site In TheExec.sites
            '--- initialize flag and get prober XY coordinate---
            CompareXcoord(site) = 0: CompareYcoord(site) = 0 ' initlize
            
            If LCase(currentJobName) Like "*cp*" Then
                CompareXcoord(site) = XCoord(site) 'from prober
                CompareYcoord(site) = YCoord(site) 'from prober
            ElseIf LCase(currentJobName) Like "*ft*" Then
                If LCase(currentJobName) Like "*wlft*" And TheExec.enableWord("STDC_ProberXY4WLFT") = True Then
                    'STDC_ProberXY4WLFT, Sampling Test Data Collection is on
                    CompareXcoord(site) = XCoord(site) 'from prober
                    CompareYcoord(site) = YCoord(site) 'from prober
                Else
                    CompareXcoord(site) = HramXCoord(site) 'from efuse
                    CompareYcoord(site) = HramYCoord(site) 'from efuse
                End If
            End If
            
            '--- i loop for searching flag info in control table ---
            TheExec.Datalog.WriteComment "-------------------------------------------------------------------"
            TheExec.Datalog.WriteComment "Site(" & site & ")" & ", (X,Y) = (" & CompareXcoord(site) & "," & CompareYcoord(site) & ")"
            
            For i = 1 To UBound(ControlTableArrayFlag)
                ComparedFlagName = vbNullString
                ComparedFlagState_Str = vbNullString
                TestResultFlagState_Str = vbNullString
                ComparedFlagState = 99
                ComparedFlagName = ControlTableArrayFlag(i, 1)
                If LCase(ControlTableArrayFlag(i, 2)) = "true" Then
                    ComparedFlagState = logicTrue
                    ComparedFlagState_Str = " True"
                ElseIf LCase(ControlTableArrayFlag(i, 2)) = "false" Then
                    ComparedFlagState = logicFalse
                    ComparedFlagState_Str = "False"
                ElseIf LCase(ControlTableArrayFlag(i, 2)) = "clear" Then
                    ComparedFlagState = logicClear
                    ComparedFlagState_Str = "Clear"
                End If
                If TheExec.sites(site).FlagState(ComparedFlagName) = logicTrue Then
                    TestResultFlagState_Str = " True"
                ElseIf TheExec.sites(site).FlagState(ComparedFlagName) = localFalse Then
                    TestResultFlagState_Str = "False"
                ElseIf TheExec.sites(site).FlagState(ComparedFlagName) = localClear Then
                    TestResultFlagState_Str = "Clear"
                End If
                If TheExec.sites(site).FlagState(ComparedFlagName) = ComparedFlagState Then ' if program flag = expected state in table
                    TheExec.sites(site).FlagState("SamplingTest_TurnOnFlag_ByFlag") = logicTrue
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & FormatNumeric(ComparedFlagName, ControlTableFlag_maxLen * -1) & ", State :" & TestResultFlagState_Str & ", Control Table :" & ComparedFlagState_Str & ", (Match up)"
                    Exit For
                Else
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & FormatNumeric(ComparedFlagName, ControlTableFlag_maxLen * -1) & ", State :" & TestResultFlagState_Str & ", Control Table :" & ComparedFlagState_Str
                End If
            Next i
            TheExec.Datalog.WriteComment "-------------------------------------------------------------------"
            
            
            If TheExec.sites(site).FlagState("SamplingTest_TurnOnFlag_ByFlag") = logicTrue Then
                ControlTableFlag_CollectedTotalDies = ControlTableFlag_CollectedTotalDies + 1
                If ControlTableFlag_CollectedTotalDies > ControlTableFlag_ExpectedTotalDies Then
                    TheExec.sites(site).FlagState("SamplingTest_TurnOnFlag_ByFlag") = logicFalse
                    TheExec.Datalog.WriteComment "Already collected over than expected count of die and turn off additional SamplingTest(By Flag)"
                    TheExec.Datalog.WriteComment "-------------------------------------------------------------------"
                    Flag_FlagControlEnable = False
                    Exit For
                End If
            End If
            
        Next site
    End If
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Parse_SubFlowControlTable() As Long
'SWLIN 2017/03/29
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Parse_SubFlowControlTable"
    Dim ControlTable As Worksheet
    Dim MaxRow_DieCount As Long
    Dim MaxRow_FlagCount As Long
    Dim CurrentRow As Long
    Dim localFlagState As Variant
    
    
    Dim ControlTableArray() As Variant
    Dim Cur_LotTmp As String
    Dim Cur_WaferIDTmp As String
    Dim CoordinateXY As String
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    Dim ws As Worksheet
    Set ws = wb.Worksheets("SamplingTest_ControlTable")
    ws.Activate
    
        If AlreadyParsed = False Then
    
          Flag_DieControlEnable = False
          Flag_FlagControlEnable = False
          
          Set ControlTable = Sheets("SamplingTest_ControlTable")
          MaxRow_DieCount = ControlTable.Cells(Rows.Count, 1).End(xlUp).Row
          MaxRow_FlagCount = ControlTable.Cells(Rows.Count, 4).End(xlUp).Row
        
          
          
          ControlTableArray = ws.range(Cells(2, 1), Cells(MaxRow_DieCount, 2)).value
          ControlTableArrayFlag = ws.range(Cells(2, 4), Cells(MaxRow_FlagCount, 6)).value
          
          If MaxRow_DieCount >= 2 Then
              ReDim ControlTableDie(MaxRow_DieCount - 2) As Dieinfo
              ControlTableDic.RemoveAll
              For CurrentRow = 1 To MaxRow_DieCount - 1
                  CoordinateXY = CStr(ControlTableArray(CurrentRow, 1)) + "," + CStr(ControlTableArray(CurrentRow, 2))
                  ControlTableDic.Add CoordinateXY, CurrentRow
              Next CurrentRow
              Flag_DieControlEnable = True
          End If
          
          
          If MaxRow_FlagCount >= 2 Then
            ControlTableFlag_maxLen = 30 'default value
            For CurrentRow = 1 To (MaxRow_FlagCount - 1)
                If ControlTableFlag_maxLen < Len(ControlTableArrayFlag(CurrentRow, 1)) Then
                    ControlTableFlag_maxLen = Len(ControlTableArrayFlag(CurrentRow, 1))
                End If
            Next CurrentRow
            ControlTableFlag_ExpectedTotalDies = ControlTableArrayFlag(1, 3) ' flag name
            Flag_FlagControlEnable = True
            Pre_LotTmp = Trim(UCase(TheExec.Datalog.Setup.LotSetup.lotId)) 'record 1st wafer lot
            Pre_WaferIDTmp = Trim(CStr(TheExec.Datalog.Setup.WaferSetup.ID)) 'record 1st wafer ID
          End If
          TheExec.Datalog.WriteComment "---------------------------------"
          TheExec.Datalog.WriteComment "Parsing Sampling Test Control Table"
          TheExec.Datalog.WriteComment "---------------------------------"
          AlreadyParsed = True

        Else 'non-1st touch
            '---------------------------------------------------------------------------------
            'For serial wafer sorting, if we sort several wafers with one program in O.I. mode,
            'We have to re-open flag of Flag-control to enable data collection sub-flow (by flag) for next wafer
            'Becasue we will turn it off when collected dies count is enough,
            '----------------------------------------------------------------------------------
            Cur_LotTmp = Trim(UCase(TheExec.Datalog.Setup.LotSetup.lotId))
            Cur_WaferIDTmp = Trim(CStr(TheExec.Datalog.Setup.WaferSetup.ID))
            If Cur_LotTmp <> Pre_LotTmp And Cur_WaferIDTmp <> Pre_WaferIDTmp Then
                ControlTableFlag_CollectedTotalDies = 0
                Flag_FlagControlEnable = True
                Pre_LotTmp = Cur_LotTmp
                Pre_WaferIDTmp = Cur_WaferIDTmp
            End If

        End If
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function SetMFP_Judgement(BasedOnCycle As Boolean, BasedOnVector As Boolean)

    If BasedOnCycle = True And BasedOnVector = True Then
        TheExec.Datalog.WriteComment "You Can Only Use One Judgement for MFP, Pattern Cycle or Pattern Vector"
        TheExec.Datalog.WriteComment "Set Vector Type as Default Setting"
        Mbist_Repair_CompareType = "Vector"
    ElseIf BasedOnCycle = False And BasedOnVector = False Then
        TheExec.Datalog.WriteComment "Please Select One Judgement for MFP, Pattern Cycle or Pattern Vector"
        TheExec.Datalog.WriteComment "Set Vector Type as Default Setting"
        Mbist_Repair_CompareType = "Vector"
    Else
        If BasedOnVector = True Then
            Mbist_Repair_CompareType = "Vector"
            TheExec.Datalog.WriteComment "Set Vector Type as Default Setting"
        ElseIf BasedOnCycle = True Then
            Mbist_Repair_CompareType = "Cycle"
            TheExec.Datalog.WriteComment "Set Cycle Type as Default Setting"
        End If
    End If
    
End Function



Public Function Create_Overlay(OverlayName As String, specName As String, SpecValue As String)
    
    Dim specName_Ary() As String
    Dim SepcValue_ary() As String
    Dim i As Long
    specName_Ary = Split(specName, ",")
    SepcValue_ary = Split(SpecValue, ",")
    
    If UBound(SepcValue_ary) <> UBound(specName_Ary) Then
        TheExec.ErrorLogMessage (vbCrLf & "Module Name: Create_Overlay" & vbCrLf & "Create Overlay Fail, Spec Value count should be same as Spec Name Count")
    Else
        With TheExec.Overlays
            If (.Contains(OverlayName) <> False) Then .Remove OverlayName
            .Add (OverlayName)
            '.Add(OverlayName).Specs.Add SpecName
        End With
        
        For i = 0 To UBound(specName_Ary)
            With TheExec.Overlays(OverlayName)
                .Specs.Add(specName_Ary(i)).value = CDbl(SepcValue_ary(i))
            End With
        Next i
    End If
    
End Function

Public Function Create_Overlay_CPM(OverlayName As String, specName As String, Optional Fail_Flag As String, Optional VoltageOffset As Double)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "Create_Overlay_CPM"
    '''**********************************************************************************'''
    '''Replace DC Spec voltage with Shmoo Lvcc voltage.                                  '''
    '''C651 Toby asked to use Shmoo Lvcc voltage for CPM Flow.                           '''
    '''So we replace voltage with Shmoo lvcc voltage                                     '''
    '''If there is a shmoo hole (value: +/-5555, +/-7777, +/-9999), binout the devices.  '''
    '''**********************************************************************************'''
    Dim i As Double
    Dim specName_Ary() As String
    Dim tempStr As String
    specName_Ary = Split(specName, ",")

    With TheExec.Overlays
        If (.Contains(OverlayName) <> False) Then .Remove OverlayName
        
        .Add (OverlayName)
        With TheExec.Overlays(OverlayName)
            For i = 0 To UBound(specName_Ary)
                .Specs.Add (specName_Ary(i))
                .Specs.item(specName_Ary(i)).value = 0
            Next i
        End With
    End With
    
    TheExec.flow.TestLimit Shmoo_Vcc_Min, , , , , scaleMilli, unitVolt, , , , , , , , , tlForceFlow
    TheExec.flow.TestLimit Shmoo_Vcc_Min.Add(VoltageOffset), , , , , scaleMilli, unitVolt, , , , , , , , , tlForceFlow
    
    TheExec.Datalog.WriteComment "********************************************"
    For Each site In TheExec.sites.Active
        tempStr = vbNullString
        If Shmoo_Vcc_Min(site) <> 0 Then
            If Abs(Shmoo_Vcc_Min(site)) = 5555 Or Abs(Shmoo_Vcc_Min(site)) = 7777 Or Abs(Shmoo_Vcc_Min(site)) = 9999 Then
                TheExec.Datalog.WriteComment "Site: " & CStr(site) & ", CPM Shmoo voltage has Shmoo hole. BinOut the devices."
                If Fail_Flag <> "" Then
                    TheExec.sites.item(site).FlagState(Fail_Flag) = logicTrue
                End If
            Else
                With TheExec.Overlays(OverlayName)
                    For i = 0 To UBound(specName_Ary)
                        .Specs.item(specName_Ary(i)).value = CDbl(Shmoo_Vcc_Min.Add(VoltageOffset))
                    Next i
                End With
                TheExec.Datalog.WriteComment "Site: " & CStr(site) & ", " & Replace(specName, "_VOP_VAR", "") & " VALT = " & CDbl(Shmoo_Vcc_Min.Add(VoltageOffset))
            End If
        Else
            TheExec.Datalog.WriteComment "Site: " & CStr(site) & ", Shmoo_Vcc_Min don't have value, please check it out."
        End If
    Next site
    TheExec.Datalog.WriteComment "********************************************"
 Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CalcCPM_FailingPatCount(blockName As String, TestCount As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "CalcCPM_FailingPatCount"
    Dim site As Variant
    Dim FailingCounts As New SiteLong
    Dim BlockResultFalg As String
    Dim itemNum As Long
    Dim temp_itemNum As String
    
        FailingCounts = 0
        For itemNum = 1 To TestCount
            'compose sflag name
            temp_itemNum = vbNullString
            temp_itemNum = right("00" & Trim(itemNum), 2)
            BlockResultFalg = "F_CPM" & temp_itemNum & "_" & blockName 'F_CPM02_CCY0
            
            For Each site In TheExec.sites
                If TheExec.sites.item(site).FlagState(BlockResultFalg) = logicFalse Then  'item pass, flag fail
                    FailingCounts = FailingCounts
                ElseIf TheExec.sites.item(site).FlagState(BlockResultFalg) = logicTrue Then 'item fail, flag pass
                    FailingCounts = FailingCounts + 1
                ElseIf TheExec.sites.item(site).FlagState(BlockResultFalg) = logicClear Then 'item no test, flag clear, treat it as fail
                     FailingCounts = FailingCounts
                End If
            Next site
        Next itemNum
    
    
    TheExec.flow.TestLimit FailingCounts, , , , , , , , , , , , , , , tlForceFlow
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
'20201214: Modified to remove the redundant argument "voltage_SelsrmBitCalc".
'20201029: Modified to use "Public Type Instance_Info".
'20201006: Modified to add the variable "is_BinSearch" to control BinSearch for cp1 and functional test for non-cp1.
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200711: Modified to use the siteDouble array "BinCut_Payload_Voltage" to store BinCut payload voltages.
'20200615: Modified to get the instance name from the argument "inst_name as string".
'20200604: Modified to remove the unused argument "powerDomain" of "Parsing_Instance_Pmode".
Public Function Create_overlay_PostBV(performance_mode As String, OverlayName As String, Pins_ForceValt As String)
    Dim SpecValue As New SiteVariant
    Dim specName As String
    Dim Pins_ForceValt_Ary() As String
    Dim i As Integer
    Dim VarStr As String
    Dim voltage_SelsrmBitCalc() As New SiteDouble '''added, 20200727
    Dim inst_info As Instance_Info
On Error GoTo errHandler
    '''init siteDouble arrary
    ReDim voltage_SelsrmBitCalc(MaxBincutPowerdomainCount)
    For i = 0 To MaxBincutPowerdomainCount
        voltage_SelsrmBitCalc(i) = 0
    Next i

    If Pins_ForceValt <> "" Then
        Dim Dic_Pins_ForceValt As New Dictionary
        Dic_Pins_ForceValt.RemoveAll
        Pins_ForceValt_Ary = Split(Pins_ForceValt, ",")
        For i = 0 To UBound(Pins_ForceValt_Ary)
            If Dic_Pins_ForceValt.Exists(Pins_ForceValt_Ary(i)) Then
                TheExec.Datalog.WriteComment "There is duplicated pin define in the Argu 'Pins_FoceValt'"
            Else
                Dic_Pins_ForceValt.Add Pins_ForceValt_Ary(i), i
            End If
        Next i
    End If
    
    '''//Initialize inst_info
    '''//Get p_mode, addi_mode, testtype, and offsettestype from test instance and its argument.
    '''20201029: Modified to use "Public Type Instance_Info".
    Call initialize_inst_info(inst_info, performance_mode)
    inst_info.is_BinSearch = False
    
    '''//Calculate BinCut payload voltages for BinCut CorePower and OtherRail.
    '''20200711: Modified to use the siteDouble array "BinCut_Payload_Voltage" to store BinCut payload voltages.
    '''20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
    '''20201006: Modified to add the variable "is_BinSearch" to control BinSearch for cp1 and functional test for non-cp1.
    '''20201029: Modified to use "Public Type Instance_Info".
    '''20201214: Modified to remove the redundant argument "voltage_SelsrmBitCalc".
    '''20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
    Call bincut_power_Setting_VT(inst_info, VBIN_RESULT(inst_info.p_mode).passBinCut, BinCut_Payload_Voltage)
    
    Dim Temp_PinAry() As String
    Dim Temp_PinStr As String
    Dim j As Double
    
    For Each site In TheExec.sites
        For i = 0 To UBound(pinGroup_BinCut)
            If UCase(pinGroup_BinCut(i)) Like UCase("*GRP*") Or UCase(pinGroup_BinCut(i)) Like UCase("*ALL*") Then
                Temp_PinStr = VddbinDomain2Pin(pinGroup_BinCut(i))
                Temp_PinAry = Split(Temp_PinStr, ",")
            Else
                ReDim Temp_PinAry(0) As String
                Temp_PinAry(0) = pinGroup_BinCut(i)
            End If
            
            For j = 0 To UBound(Temp_PinAry)
                If Dic_Pins_ForceValt.Exists(Temp_PinAry(j)) Then
                    VarStr = "_VOP_VAR"
                Else
                    VarStr = "_VAR"
                End If
                
                If UCase(Temp_PinAry(j)) Like "*_FT" Then Temp_PinAry(j) = Replace(Temp_PinAry(j), "_FT", "", compare:=vbTextCompare)
                
                If i = 0 Then
                    SpecValue = BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_BinCut(i))) * 0.001
                    specName = Temp_PinAry(j) & VarStr
                Else
                    SpecValue = SpecValue & "," & BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_BinCut(i))) * 0.001
                    specName = specName & "," & Temp_PinAry(j) & VarStr
                End If
            Next j
        Next i
    Next site
    'specName = Join(pinGroup_BinCut, "_VOP_VAR,") & "_VOP_VAR"
           
    Call Create_Overlay_MultiSite(OverlayName, specName, SpecValue)
Exit Function
errHandler:
   TheExec.Datalog.WriteComment "Error encountered in VBT Function of RestoreSkipTestBin2Site"
   If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Create_Overlay_MultiSite(OverlayName As String, specName As String, SpecValue As SiteVariant)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "Create_Overlay_MultiSite"
    
    Dim i As Double
    Dim specName_Ary() As String
    Dim specValue_Ary() As String
    
    specName_Ary = Split(specName, ",")

    With TheExec.Overlays
        If (.Contains(OverlayName) <> False) Then .Remove OverlayName
        
        .Add (OverlayName)
        With TheExec.Overlays(OverlayName)
            For i = 0 To UBound(specName_Ary)
                .Specs.Add (specName_Ary(i))
                .Specs.item(specName_Ary(i)).value = 0
            Next i
        End With
    End With
    
'    theexec.Datalog.WriteComment "********************************************"
    TheExec.Datalog.WriteComment "Overlay Name : " & OverlayName
'    theexec.Datalog.WriteComment "********************************************"
    
    For Each site In TheExec.sites.Active
        specValue_Ary = Split(CStr(SpecValue(site)), ",")
        
        With TheExec.Overlays(OverlayName)
            For i = 0 To UBound(specName_Ary)
                If dictPin2Dcspec.Exists(specName_Ary(i)) Then
                    .Specs.item(specName_Ary(i)).value = CDbl(specValue_Ary(i))
'                    If InStr(1, specName_Ary(i), "_VOP_VAR") = 0 Then
'                        theexec.Datalog.WriteComment "Site: " & CStr(site) & ", " & Replace(specName_Ary(i), "_VAR", "") & " VMAIN = " & CDbl(specValue_Ary(i)) & " V"
'                    Else
'                        theexec.Datalog.WriteComment "Site: " & CStr(site) & ", " & Replace(specName_Ary(i), "_VOP_VAR", "") & " VALT = " & CDbl(specValue_Ary(i)) & " V"
'                    End If
                Else
''                    specName_Ary(i) = Replace(specName_Ary(i), "_VOP", "")
''                    If dictPin2Dcspec.Exists(specName_Ary(i)) Then
''                        .Specs.Item(specName_Ary(i)).Value = CDbl(specValue_Ary(i))
''                        TheExec.Datalog.WriteComment "Site: " & CStr(site) & ", " & Replace(specName_Ary(i), "_VAR", "") & " VMAIN = " & CDbl(specValue_Ary(i)) & " V"
''                    Else
                        TheExec.Datalog.WriteComment "Can Not Find Spec Name : " & specName_Ary(i) & ", please check"
                        GoTo errHandler
''                    End If
                End If
            Next i
        End With
        
'        theexec.Datalog.WriteComment "********************************************"
    Next site
    
 Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ClearOverlay(OverlayName As String)
Dim OverlayName_ary() As String
Dim i As Double
OverlayName_ary = Split(OverlayName, ",")

    For i = 0 To UBound(OverlayName_ary)
        With TheExec.Overlays
            If (.Contains(OverlayName_ary(i)) <> False) Then
                .Remove OverlayName_ary(i)
                TheExec.Datalog.WriteComment "Remove Overlay Name: " & OverlayName_ary(i)
            End If
        End With
    Next i

End Function

Public Function SetFuse_baseOnFlowVariant(FlowVariantName As String, FuseBank As String, FuseCate As String)
On Error GoTo errHandler
Dim funcName As String:: funcName = "SetFuse_baseOnFlowVariant"
Dim site As Variant
Dim fuseValue As New SiteDouble
Dim FlowVarAry() As String
Dim bitStr As String
Dim var As Variant

    FlowVarAry = Split(FlowVariantName, ",")     '''''210625 MHV/MLV index to array
    
    For Each site In TheExec.sites
        bitStr = vbNullString
        For Each var In FlowVarAry
            bitStr = CStr(TheExec.sites.item(site).SiteVariableValue(var)) + bitStr     '''''210625 convert from flow variable to string
        Next
        fuseValue(site) = GlbUtility.Bin2Dec(CStr(bitStr))     '''''210625 convert from flow variable to string
        TheExec.Datalog.WriteComment vbTab & "Site: " & site & ", Index: " & FlowVariantName & ", Value = " & fuseValue(site)
    Next site
    
    'Call auto_eFuse_SetWriteVariable_SiteAware(FuseBank, FuseCate, fuseValue, True)
    
    ''====20201230 add for efuse new code====
    Dim opbank As eFuseBdfBank ''20201112 add for obj
    Dim field As eFuseBdfField
    Set opbank = GetBdfBank(FuseBank)
    Set field = opbank.Fields(FuseCate)
    opbank.SetEfuse field.name, fuseValue, , , , , True
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
