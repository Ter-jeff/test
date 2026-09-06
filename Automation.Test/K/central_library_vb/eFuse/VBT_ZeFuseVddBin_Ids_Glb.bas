Attribute VB_Name = "VBT_ZeFuseVddBin_Ids_Glb"
Option Explicit
#Const ApplyBinCut = True

'202004xx disable for AP-----
#If Not ApplyBinCut Then
    Public CurrentPassBinCutNum As New SiteLong
    ''202004xx for ap
    'Enum VBin_Performance_mode
    Public Const VDD_ERROR = 60
    'End Enum
    Public Function VddBinStr2Enum(VddBinName As String) As VBin_Performance_mode
    End Function
#End If

Public Sub ReadEfuseDataFromBinCut(Optional getLimitsOnly As Boolean = False)
On Error GoTo errHandler

    Call GetIdsValues(getLimitsOnly)
    Call GetBinCutValuesAndLimits(getLimitsOnly)
    'Call Ze_Read_DVFM_To_GradeVDD
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "ReadEfuseDataFromBinCut")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
'202004xx disable for AP end-----

Public Sub GetIdsValues(Optional getLimitsOnly As Boolean = False) 'eFuse stored ids values
On Error GoTo errHandler
Dim bankstr As Variant, bank As eFuseBdfBank
Dim fieldStr As Variant, field As eFuseBdfField
Dim m_Pmode As Long, values As New SiteVariant, PatPF As New SiteBoolean
Dim site As Variant

    PatPF = True
    For Each bankstr In BdfDataBase.Banks.Keys
            If bankstr <> Empty Then
                Set bank = BdfDataBase.Banks(bankstr)
                If bank.HadIdsFuse Then
                           For Each fieldStr In bank.Fields.Keys
                                   Set field = bank.Fields(fieldStr)
                                   If field.Algorithm = alg_ids Then
                                   For Each site In TheExec.sites
                                        If Not getLimitsOnly Then field.FuseMeasureValue = field.IdsValue_A
'#If ApplyBinCut Then
'                                        If GlbUtility.IsOnline Then field.Hlimit = Get_IDSMax_fromBinCut(field.Name) ' offline, skip get it from BinCut sheet
'#End If
                                   Next
                                   End If
                           Next
                End If
            End If
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetIdsValues")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

''for offline
Public Sub PutIdsCodes()
On Error GoTo errHandler
Dim bankstr As Variant, bank As eFuseBdfBank
Dim fieldStr As Variant, field As eFuseBdfField
Dim m_Pmode As Long, values As New SiteVariant, PatPF As New SiteBoolean

    PatPF = True
    For Each bankstr In BdfDataBase.Banks.Keys
        If bankstr <> Empty Then
            Set bank = BdfDataBase.Banks(bankstr)
            If bank.HadIdsFuse Then
                For Each fieldStr In bank.DicIds ' This dictionary contains items for current test stage only
                    Set field = bank.Fields(fieldStr)
                    If field.DefaultOrReal = dr_real Then
                        Set values = ObtainIdsValues(field)
                        bank.SetEfuse field.name, values, PatPF
                    End If
                Next
            End If
        End If
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "PutIdsCodes")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function ObtainIdsValues(field As eFuseBdfField) As SiteVariant
On Error GoTo errHandler
Dim funcName As String: funcName = "ObtainIdsValues"
Dim result As New SiteDouble, fieldStr As String: fieldStr = LCase(field.name)
Dim xPowerPin As String
Dim site As Variant
Dim VddField As New eFuseBdfField
Dim idsField As New eFuseBdfField
Dim opbank As New eFuseBdfBank
Dim vddfieldstr As Variant
Dim m_Pmode As Long
Dim VddIdsmatch As Boolean: VddIdsmatch = False
Dim fieldStage As String

    If All_Power_data_IDS_GB Is Nothing Then
        GlbUtility.MessageBox "All_Power_data_IDS_GB is nothing, must run IDS firstly"
        GoTo errHandler
    End If
    xPowerPin = Replace(field.name, "ids_", "")
    '202004xx for ap
    'Result = All_Power_data_IDS_GB.Pins(xPowerPin)
    If Not GlbUtility.OnlineMode Then

        Set opbank = GetBdfBank(field.bankName)

        fieldStage = BdfDataBase.GetRealStage(field.BlowLocation)
        fieldStage = Replace(UCase(fieldStage), "_EARLY", "")
        If Not (opbank.CreateFakeValueTime = 0 And (Not PseudoFuseEnable)) And Not (fieldStage = GlbUtility.currStage) Then
            For Each site In TheExec.sites
                result = 0
            Next site
            Set ObtainIdsValues = result
            Exit Function
        End If

        For Each vddfieldstr In opbank.Fields 'This dictionary contains items for current test stage only
            Set VddField = opbank.Fields(vddfieldstr)
            fieldStage = BdfDataBase.GetRealStage(VddField.BlowLocation)
            fieldStage = Replace(UCase(fieldStage), "_EARLY", "")

            If VddField.Algorithm = alg_vddbin And (fieldStage = GlbUtility.currStage Or GlbUtility.testedStages.Exists(fieldStage)) Then
                m_Pmode = VddBinStr2Enum(Replace(LCase(VddField.name), "_shadow", ""))
                Set idsField = GetMapIdsField(AllBinCut(m_Pmode).IDS_MAPPING)
                
                If idsField.name = field.name Then
                    VddIdsmatch = True
                    For Each site In TheExec.sites
                        result = Round(((AllBinCut(m_Pmode).IDS_CP_LIMIT * 0.001) - idsField.Llimit) * Rnd + idsField.Llimit, 5)
                    Next site
                    Exit For
                End If
            End If
        Next
        
        If Not VddIdsmatch Then
            For Each site In TheExec.sites
                result = Round((field.Hlimit - field.Llimit) * Rnd() + field.Llimit, 5)  'to A 'Pseudo ids in offine
            Next
        End If
    End If
    Set ObtainIdsValues = result

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "ObtainIdsValues")
    If AbortTest Then Exit Function Else Resume Next
End Function

'202004xx disable for AP-----
Public Sub GetBinCutValuesAndLimits(Optional getLimitsOnly As Boolean = False) 'From DSSC result
On Error GoTo errHandler
#If ApplyBinCut Then
Dim bankstr As Variant, bank As eFuseBdfBank
Dim fieldStr As Variant, field As eFuseBdfField
Dim m_Pmode As Long, values As New SiteVariant, PatPF As New SiteBoolean
Dim field_normal As eFuseBdfField, field_addition As eFuseBdfField
Dim need_calculate As New SiteBoolean: need_calculate = True
Dim need_calculate_addition As New SiteBoolean: need_calculate_addition = True
Dim site As Variant
Dim fieldStage As String

    PatPF = True
    If BdfDataBase.Banks("CFG").Fields.Exists("product_identifier") Then
        Set field_normal = BdfDataBase.Banks("CFG").Fields("product_identifier")
        For Each site In TheExec.sites
            CurrentPassBinCutNum_normal(site) = field_normal.DsscDecValue + 1
            If CurrentPassBinCutNum_normal(site) > Total_Bincut_Num Then
                need_calculate(site) = False
                Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetBinCutValuesAndLimits", "site:" & site & ", product_identifier " & CurrentPassBinCutNum_normal(site) & " > Total_Bincut_Num " & Total_Bincut_Num)
                TheExec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="product_identifier Error"
            End If
        Next site
        CurrentPassBinCutNum_additional = CurrentPassBinCutNum_normal
    Else
        CurrentPassBinCutNum_normal = 1
    End If
    
    If BdfDataBase.Banks("CFG").Fields.Exists("product_identifier_shadow") Then
        Set field_addition = BdfDataBase.Banks("CFG").Fields("product_identifier_shadow")
        If field_addition.DefaultOrReal = dr_real Then
            For Each site In TheExec.sites
                CurrentPassBinCutNum_additional(site) = field_addition.DsscDecValue + 1
                If CurrentPassBinCutNum_additional(site) > Total_Bincut_Num Then
                    need_calculate_addition(site) = False
                    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetBinCutValuesAndLimits", "site:" & site & ", product_identifier_shadow " & CurrentPassBinCutNum_additional(site) & " > Total_Bincut_Num " & Total_Bincut_Num)
                    TheExec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="product_identifier_shadow Error"
                End If
            Next site
        End If
    End If

    For Each bankstr In BdfDataBase.Banks.Keys
        If bankstr <> Empty Then
            Set bank = BdfDataBase.Banks(bankstr)
            If bank.HadVddBinFuse Then
                For Each fieldStr In bank.Fields.Keys
                    Set field = bank.Fields(fieldStr)
                    If field.Algorithm = alg_vddbin Then
                        fieldStage = BdfDataBase.GetRealStage(field.BlowLocation)
                        fieldStage = Replace(UCase(fieldStage), "_EARLY", "")
                        For Each site In TheExec.sites
                            If Not getLimitsOnly Then field.FuseMeasureValue = field.VddBinValue_V
                        Next
                        If (fieldStage = GlbUtility.currStage) Or GlbUtility.testedStages.Exists(fieldStage) Then
                            m_Pmode = VddBinStr2Enum(Replace(LCase(field.name), "_shadow", ""))
                            For Each site In TheExec.sites
                                If field.DefaultOrReal = dr_real Then
                                    Dim MaxLevelIndex As Long
                                                  
                                    If LCase(field.BlowLocation) <> LCase(BincutAdditionalSheetName) Then
                                        If need_calculate(site) Then
                                            MaxLevelIndex = EfuseBinCut(m_Pmode, CurrentPassBinCutNum_normal).Mode_Step
                                            field.BVLLimit = EfuseBinCut(m_Pmode, CurrentPassBinCutNum_normal).CP_Vmin(MaxLevelIndex) + EfuseBinCut(m_Pmode, CurrentPassBinCutNum_normal).CP_GB(MaxLevelIndex)
                                            field.BVHLimit = EfuseBinCut(m_Pmode, CurrentPassBinCutNum_normal).CP_Vmax(0) + EfuseBinCut(m_Pmode, CurrentPassBinCutNum_normal).CP_GB(0)
                                        End If
                                    Else
                                        If need_calculate_addition(site) Then
                                            MaxLevelIndex = EfuseBinCutAddition(m_Pmode, CurrentPassBinCutNum_additional).Mode_Step
                                            field.BVLLimit = EfuseBinCutAddition(m_Pmode, CurrentPassBinCutNum_additional).CP_Vmin(MaxLevelIndex) + EfuseBinCutAddition(m_Pmode, CurrentPassBinCutNum_additional).CP_GB(MaxLevelIndex)
                                            field.BVHLimit = EfuseBinCutAddition(m_Pmode, CurrentPassBinCutNum_additional).CP_Vmax(0) + EfuseBinCutAddition(m_Pmode, CurrentPassBinCutNum_additional).CP_GB(0)
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="GetBinCutValuesAndLimits"
                            Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetBinCutValuesAndLimits", "BincutDict_Not_Exist! " & bankstr & ">>" & fieldStr)
                        End If
                    End If
                Next
            End If
        End If
    Next
#End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetBinCutValuesAndLimits")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
'202004xx disable for AP end-----

''202004xx disable for AP-----
Public Function Ze_Read_DVFM_To_GradeVDD() As Long
On Error GoTo errHandler
#If ApplyBinCut Then
Dim p_mode As Long
Dim fieldStr As String
Dim bank As String, opbank As eFuseBdfBank, field As eFuseBdfField, gb As Double, stage As String
Dim site As Variant

    Call CheckeFusePassBinCutNum
    stage = LCase(GlbUtility.currStage)
    For p_mode = 0 To MaxPerformanceModeCount - 1
        If AllBinCut(p_mode).Used = True Then
            fieldStr = VddBinName(p_mode)
            bank = BdfDataBase.DicVddBinBankMap(fieldStr)
            Set opbank = GetBdfBank(bank)
            Set field = opbank.Fields(fieldStr)
            For Each site In TheExec.sites
                If stage Like "*cp2*" Then
                    gb = BinCut(p_mode, CurrentPassBinCutNum).CP2_GB(0)
                ElseIf stage Like "*qa*" Then
                    gb = BinCut(p_mode, CurrentPassBinCutNum).FTQA_GB(0)
                ElseIf stage Like "*ft1*" Or stage Like "*wlft*" Then
                    gb = BinCut(p_mode, CurrentPassBinCutNum).FT1_GB(0)
                ElseIf stage Like "*ft2*" Then
                    gb = BinCut(p_mode, CurrentPassBinCutNum).FT2_GB(0)
                Else
                    GlbUtility.MessageBox "The Job selection Error!! at ""Ze_Read_DVFM_To_GradeVDD"""
                End If
                
                VBIN_RESULT(p_mode).GRADEVDD = field.FuseMeasureValue
                VBIN_RESULT(p_mode).GRADE = VBIN_RESULT(p_mode).GRADEVDD - gb
                VBIN_RESULT(p_mode).passBinCut = CurrentPassBinCutNum
            Next site
        End If
    Next p_mode
#End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "Ze_Read_DVFM_To_GradeVDD")
    If AbortTest Then Exit Function Else Resume Next
End Function
''202004xx disable for AP end-----

'202004xx disable for AP-----
Private Function Get_IDSMax_fromBinCut(CorePwrofIds As String) As Double
#If ApplyBinCut Then
On Error GoTo errHandler
Dim xPowerPin As String, MaxIDSValue As Double, i As Integer, corePwrList As String

    corePwrList = "vdd_pcpu$|vdd_ecpu$|vdd_gpu$|vdd_soc$|vdd_dcs" '_ddr
    xPowerPin = Replace(UCase(CorePwrofIds), "IDS_", "")
    xPowerPin = LCase(xPowerPin)
    MaxIDSValue = 0 ' search all power and record maximum ids spec
    
    If Not xPowerPin Like "*_85" Then  'room temp.
        If xPowerPin Like "*vdd_fixed" Or xPowerPin Like "*vdd_low" Then xPowerPin = xPowerPin & "_grp"
        If GlbUtility.IsStrMatch(xPowerPin, corePwrList) Then ' for core power only
            For i = 0 To UBound(AllBinCut())
                If AllBinCut(i).Used = True Then
                    If UCase(AllBinCut(i).powerPin) = UCase(xPowerPin) Then
                        If MaxIDSValue < AllBinCut(i).IDS_CP_LIMIT Then MaxIDSValue = AllBinCut(i).IDS_CP_LIMIT
                    End If
                End If
            Next i
        Else
            xPowerPin = UCase(xPowerPin) ' for fixed, sram, low
            MaxIDSValue = AllBinCut(VddBinStr2Enum(xPowerPin)).IDS_CP_LIMIT
        End If
    Else 'high temp.
        xPowerPin = Replace(xPowerPin, "_85", "")
        If xPowerPin Like "*vdd_fixed" Or xPowerPin Like "*vdd_low" Then xPowerPin = xPowerPin & "_grp"
        If GlbUtility.IsStrMatch(xPowerPin, corePwrList) Then ' for core power only
            For i = 0 To UBound(AllBinCut())
                If AllBinCut(i).Used = True Then
                    If UCase(AllBinCut(i).powerPin) = UCase(xPowerPin) Then
                        If MaxIDSValue < AllBinCut(i).IDS_FT_LIMIT Then MaxIDSValue = AllBinCut(i).IDS_FT_LIMIT
                    End If
                End If
            Next i
        Else
            xPowerPin = UCase(xPowerPin) ' for fixed, sram, low
            MaxIDSValue = AllBinCut(VddBinStr2Enum(xPowerPin)).IDS_FT_LIMIT
        End If
    End If
    Get_IDSMax_fromBinCut = MaxIDSValue * 0.001 'mA to A

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "Get_IDSMax_fromBinCut")
    If AbortTest Then Exit Function Else Resume Next
#End If
End Function

Public Sub PutBinCutCodes()
On Error GoTo errHandler
#If ApplyBinCut Then
Dim bankstr As Variant, bank As eFuseBdfBank
Dim fieldStr As Variant, field As eFuseBdfField
Dim m_Pmode As Long, values As New SiteVariant, PatPF As New SiteBoolean

    PatPF = True
    For Each bankstr In BdfDataBase.Banks.Keys
       If bankstr <> Empty Then
                Set bank = BdfDataBase.Banks(bankstr)
                If bank.HadVddBinFuse Then
                           For Each fieldStr In bank.DicBinCut ' This dictionary contains items for current test stage only
                                   Set field = bank.Fields(fieldStr)
                                   If field.DefaultOrReal = dr_real Then
                                        m_Pmode = VddBinStr2Enum(field.name)
                                        '202004xx for ap
                                        'If m_Pmode <> VDD_ERROR Then
                                             Set values = GenVddBinFuseCode(field.name, m_Pmode, field)
                                             bank.SetEfuse field.name, values, PatPF
                                        'Else
                                        'GlbUtility.WriteErrDlg "Missing VddBin Value for """ & field.Name & """"
                                        'End If
                                  End If
                           Next
                End If
        End If
    Next
#End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "PutBinCutCodes")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function GenVddBinFuseCode(mode As String, pmode As Long, field As eFuseBdfField) As SiteVariant
On Error GoTo errHandler
#If ApplyBinCut Then
Dim vbin As Double, result As New SiteVariant, Base As Double
Dim site As Variant

    Base = BdfDataBase.BaseVoltage
    For Each site In TheExec.sites
        vbin = VBIN_RESULT(pmode).GRADEVDD(site)
        If vbin = -1 Then vbin = 0
        field.MeasureValue = vbin
        If Not GlbUtility.OnlineMode Then vbin = GetOfflineVddBin(pmode, field) 'Base + 240 * Rnd(Now)
        'If Not GlbUtility.IsOnline Then vbin = GetOfflineVddBin(pmode, field) 'Base + 240 * Rnd(Now)
        If vbin = 0 Then
            result = 0
        Else
            result = GlbUtility.CeilingValue((vbin - Base) / field.Resolution, 1)
        End If
        If eFusePrinted Then
            GlbUtility.WriteDlg "Site(" & site & ") " & mode & "=" & vbin & "mV"
        End If
    Next
    Set GenVddBinFuseCode = result
#End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GenVddBinFuseCode")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetOfflineVddBin(pmode As Long, field As eFuseBdfField) As Double
On Error GoTo errHandler
#If ApplyBinCut Then
Dim ids_current As Double, PASSBIN As Long: PASSBIN = 1
Dim k As Long, cal_voltage As Double, final_voltage As Double, Remainder As Double, idsField As eFuseBdfField: k = 0
Dim fieldname_mod As String
Dim i As Long

    fieldname_mod = Replace(LCase(field.name), "_shadow", "")

    Set idsField = BdfDataBase.DicVddBinIdsFieldMap(fieldname_mod)

    ids_current = idsField.MeasureValue * 0.001

    If LCase(field.BlowLocation) = LCase(BincutAdditionalSheetName) Then
        cal_voltage = EfuseBinCutAddition(pmode, PASSBIN).c(k) - EfuseBinCutAddition(pmode, PASSBIN).m(k) * (log(ids_current * 1000) / log(10))
        Remainder = cal_voltage / field.Resolution
        Remainder = Floor(Remainder)
        cal_voltage = Remainder * field.Resolution
        If cal_voltage > EfuseBinCutAddition(pmode, PASSBIN).CP_Vmax(k) Then
              final_voltage = EfuseBinCutAddition(pmode, PASSBIN).CP_Vmax(k) + EfuseBinCutAddition(pmode, PASSBIN).CP_GB(k)
        Else
              If cal_voltage < EfuseBinCutAddition(pmode, PASSBIN).CP_Vmin(k) Then
                  final_voltage = EfuseBinCutAddition(pmode, PASSBIN).CP_Vmin(k) + EfuseBinCutAddition(pmode, PASSBIN).CP_GB(k)
              Else
                  final_voltage = cal_voltage + EfuseBinCutAddition(pmode, PASSBIN).CP_GB(k)
              End If
        End If
    Else
        cal_voltage = EfuseBinCut(pmode, PASSBIN).c(k) - EfuseBinCut(pmode, PASSBIN).m(k) * (log(ids_current * 1000) / log(10))
        Remainder = cal_voltage / field.Resolution
        Remainder = Floor(Remainder)
        cal_voltage = Remainder * field.Resolution
        If cal_voltage > EfuseBinCut(pmode, PASSBIN).CP_Vmax(k) Then
              final_voltage = EfuseBinCut(pmode, PASSBIN).CP_Vmax(k) + EfuseBinCut(pmode, PASSBIN).CP_GB(k)
        Else
              If cal_voltage < BinCut(pmode, PASSBIN).CP_Vmin(k) Then
                  final_voltage = EfuseBinCut(pmode, PASSBIN).CP_Vmin(k) + EfuseBinCut(pmode, PASSBIN).CP_GB(k)
              Else
                  final_voltage = cal_voltage + EfuseBinCut(pmode, PASSBIN).CP_GB(k)
              End If
        End If
    End If
    
    field.MeasureValue = final_voltage
    GetOfflineVddBin = final_voltage
    If eFusePrinted Then
        GlbUtility.WriteDlg "bincut simulation =>" & idsField.name & " = " & ids_current & " A ===> vddbin = " & GetOfflineVddBin
    End If
#End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetOfflineVddBin")
    If AbortTest Then Exit Function Else Resume Next
End Function

'Disscuss to remove for future project
'Public Function Ids_BinCut_PreCheck() As Long
'#If ApplyBinCut Then
'On Error GoTo errHandler
'Dim funcName As String: funcName = "Ids_BinCut_PreCheck"
'Dim BinCutBank As Long: BinCutBank = 0
'Dim bankstr As Variant, bank As eFuseBdfBank, fieldStr As Variant, field As eFuseBdfField
'                    For Each bankstr In BdfDataBase.Banks.Keys
'                                Set bank = GetBdfBank(CStr(bankstr))
'                                If bank.HadVddBinFuse Then
'                                     BinCutBank = BinCutBank + 1
'                                End If
'                    Next
'If BinCutBank <= 1 Then
'        TheExec.Flow.TestLimit resultVal:=BinCutBank, lowVal:=BinCutBank, hiVal:=BinCutBank, Tname:="eFuse_noBinCut"
'        Exit Function
'End If
'
'Dim value As New SiteDouble, sumAllBanksIds As New SiteDouble, sumAllBanksVddbin As New SiteDouble
'Dim zeroIdsDetect As New SiteBoolean, zeroVddbinDetect As New SiteBoolean
'sumAllBanksIds = 0: sumAllBanksVddbin = 0
'For Each bankstr In BdfDataBase.Banks.Keys
'        Set bank = GetBdfBank(CStr(bankstr))
'              For Each fieldStr In bank.Fields.Keys
'                      Set field = bank.Fields(fieldStr)
'                      If field.Algorithm = alg_vddbin Then
'                               For Each site In TheExec.sites
'                                    value = field.DsscDecValue
'                                    If value = 0 Then zeroVddbinDetect = True
'                                    bank.SumVddbinCodes = bank.SumVddbinCodes.Add(value)
'                                    sumAllBanksVddbin = sumAllBanksVddbin.Add(value)
'                               Next
'                      ElseIf field.Algorithm = alg_ids And LCase(field.BlowLocation) = "cp1" Then
'                               For Each site In TheExec.sites
'                                    value = field.DsscDecValue
'                                    If value = 0 Then zeroIdsDetect = True
'                                    bank.SumIdsCodes = bank.SumIdsCodes.Add(value)
'                                    sumAllBanksIds = sumAllBanksIds.Add(value)
'                               Next
'                      End If
'              Next
'Next
'
'    For Each site In TheExec.sites
'        If (sumAllBanksIds + sumAllBanksVddbin) > 0 And (zeroIdsDetect Or zeroVddbinDetect) Then
'                For Each bankstr In BdfDataBase.Banks.Keys
'                       Set bank = GetBdfBank(CStr(bankstr))
'                       value = value + bank.SumIdsCodes + bank.SumVddbinCodes
'                       For Each fieldStr In bank.Fields.Keys
'                              Set field = bank.Fields(fieldStr)
'                              If field.Algorithm = alg_vddbin And field.DsscDecValue = 0 Then
'                                        TheExec.Datalog.WriteComment vbTab & "<WARNING>" & bank.name & "-[" & field.name & "]BinCut:: There is the Empty (Zero) case."
'                              ElseIf field.Algorithm = alg_ids And LCase(field.BlowLocation) = "cp1" And field.DsscDecValue = 0 Then
'                                        TheExec.Datalog.WriteComment vbTab & "<WARNING>" & bank.name & "-[" & field.name & "]IDS:: There is the Empty (Zero) case."
'                              End If
'                       Next
'                Next
'        Else
'            value = 0 'set it Pass
'        End If
'    Next
'    TheExec.Flow.TestLimit resultVal:=value, lowVal:=0, hiVal:=0
'Exit Function
'errHandler:
'    GlbUtility.WriteDlg "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'#End If
'End Function

Public Function Ids_BinCut_PostCheck() As Long
#If ApplyBinCut Then
On Error GoTo errHandler
Dim funcName As String: funcName = "Ids_BinCut_PostCheck"
Dim m_EQNum As New SiteLong, maxEQ As New SiteLong, GRADEVDD As New SiteDouble
Dim bankstr As Variant, bank As eFuseBdfBank, fieldStr As Variant, field As eFuseBdfField, site As Variant
    'Call CheckeFusePassBinCutNum
    For Each bankstr In BdfDataBase.Banks.Keys
        If bankstr <> Empty Then
            Set bank = BdfDataBase.Banks(bankstr)
            If bank.HadVddBinFuse Then
                For Each fieldStr In bank.Fields.Keys
                    Set field = bank.Fields(fieldStr)
                    If field.Algorithm = alg_vddbin And (field.BlowLocation = GlbUtility.currStage Or GlbUtility.testedStages.Exists(field.BlowLocation)) Then
                        If Not BdfDataBase.DicVddBinIdsFieldMap.Exists(Replace(LCase(fieldStr), "_shadow", "")) Then
                            GlbUtility.WriteErrDlg "Can't get ids mapping field! " & bankstr & ">>" & fieldStr
                        Else
                            For Each site In TheExec.sites
                                If LCase(field.BlowLocation) = LCase(BincutAdditionalSheetName) Then
                                    m_EQNum = GetVddBinEqu(field, CurrentPassBinCutNum_normal, maxEQ, GRADEVDD, True)
                                    TheExec.Flow.TestLimit resultVal:=m_EQNum, lowVal:=1, hiVal:=maxEQ, Tname:=field.name & "_EQ"
                                Else
                                    m_EQNum = GetVddBinEqu(field, CurrentPassBinCutNum_normal, maxEQ, GRADEVDD)
                                    TheExec.Flow.TestLimit resultVal:=m_EQNum, lowVal:=1, hiVal:=maxEQ, Tname:=field.name & "_EQ"
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "Ids_BinCut_PostCheck")
    If AbortTest Then Exit Function Else Resume Next
#End If
End Function

Public Function GetVddBinEqu(field As eFuseBdfField, m_BinCutNum As SiteLong, ByRef maxEQ As SiteLong, ByRef GRADEVDD As SiteDouble, Optional isAddition As Boolean = False, Optional showPrint As Boolean = True, Optional useFusedResult As Boolean = True) As Long
On Error GoTo errHandler
#If ApplyBinCut Then
Dim m_idsCurrent As New SiteDouble, m_GRADEVDD As New SiteDouble
Dim m_Pmode As Long, m_EQNum As Long, m_totalEQStepNum As Long
Dim idsField As eFuseBdfField, result As Long
Dim fieldname_mod As String
Dim i As Long
    
    result = 999
    fieldname_mod = Replace(LCase(field.name), "_shadow", "")
    
    If BdfDataBase.DicVddBinIdsFieldMap.Exists(fieldname_mod) Then
        Set idsField = BdfDataBase.DicVddBinIdsFieldMap(fieldname_mod)
                        
        If Not GlbUtility.OnlineMode Then
            m_Pmode = BdfDataBase.DicVddBinPmodeMap(fieldname_mod)
        Else
            m_Pmode = VddBinStr2Enum(fieldname_mod)
        End If

        '20220629, modify for preset&check flow, use all same as bincut
        If useFusedResult = True Then
            m_idsCurrent = idsField.FuseMeasureValue
            m_GRADEVDD = field.FuseMeasureValue
        Else
            'preset&check flow, use all same as bincut
            If idsField.FuseMeasureValue <> 0 Then
                m_idsCurrent = idsField.FuseMeasureValue
            ElseIf field.TrimAteDecValue * field.Resolution <> 0 Then
                m_idsCurrent = (GlbUtility.Hex2Dbl(idsField.TrimAteValue) * idsField.Resolution) * 0.001
            End If
            
            m_GRADEVDD = GlbUtility.Hex2Dbl(field.TrimAteValue) * field.Resolution + BdfDataBase.BaseVoltage
        End If
                           
        GRADEVDD = m_GRADEVDD * 0.001
        
        If isAddition = True Then
            m_totalEQStepNum = EfuseBinCutAddition(m_Pmode, CLng(m_BinCutNum)).Mode_Step + 1
        Else
            m_totalEQStepNum = EfuseBinCut(m_Pmode, CLng(m_BinCutNum)).Mode_Step + 1
        End If
        
        maxEQ = m_totalEQStepNum
        result = Get_EQN_Voltage_Per_Site(CDbl(m_idsCurrent), m_Pmode, CDbl(m_GRADEVDD), field.Resolution, CLng(m_BinCutNum), showPrint, isAddition)
        If result = 999 Then
            GRADEVDD = 0
            GlbUtility.WriteErrDlg "The EQN Number Can Not Be Found!!"
        End If
    Else
        GlbUtility.WriteErrDlg "<FATAL> This eFuse field (" + field.name + ") didn't get any mapping of IDS field!"
    End If
    GetVddBinEqu = result
#End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetVddBinEqu")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub GetVddBinPmodeMap()
On Error GoTo errHandler
#If ApplyBinCut Then
Dim bank As eFuseBdfBank, bankstr As Variant
Dim fieldStr As Variant, field As eFuseBdfField, m_Pmode As Long, idsField As eFuseBdfField
Dim vddName As String: vddName = vbNullString

    If BdfDataBase.DicVddBinPmodeMap.Count <> 0 Then
        GlbUtility.DicCleaner BdfDataBase.DicVddBinPmodeMap
        GlbUtility.DicCleaner BdfDataBase.DicVddBinIdsFieldMap
        GlbUtility.DicCleaner BdfDataBase.DicVddBinBankMap
        Exit Sub
    End If
    For Each bankstr In BdfDataBase.Banks.Keys
        Set bank = BdfDataBase.Banks(bankstr)
        For Each fieldStr In bank.Fields
            Set field = bank.Fields(fieldStr)
                If field.Algorithm = alg_vddbin And _
                (TheExec.CurrentJob = UCase(field.BlowLocation) Or _
                GlbUtility.testedStages.Exists(UCase(field.BlowLocation))) Then
                    vddName = Replace(CStr(fieldStr), "_shadow", "")
                    If Not VddbinPmodeDict.Exists(UCase(vddName)) Then
                        TheExec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="BincutDict_Not_Exist"
                        Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetVddBinPmodeMap", fieldStr & " doesn't have the matched definition of Enum p_mode in VddBinStr2Enum")
                    Else
                        m_Pmode = VddBinStr2Enum(vddName)
                        If (Not BdfDataBase.DicVddBinPmodeMap.Exists(vddName)) And (Not BdfDataBase.DicVddBinBankMap.Exists(vddName)) Then
                            BdfDataBase.DicVddBinPmodeMap.Add LCase(vddName), m_Pmode
                            BdfDataBase.DicVddBinBankMap.Add LCase(vddName), bankstr
                        Else
                            ''do nothing
                        End If

                        If Not Flag_IDS_Mapping_enable Then
                            Set idsField = GetMapIdsField(AllBinCut(m_Pmode).powerPin)
                        Else
                            Set idsField = GetMapIdsField(AllBinCut(m_Pmode).IDS_MAPPING)
                        End If

                        If idsField Is Nothing Then
                            Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetVddBinPmodeMap", "Can't get ids mapping field! " & bankstr & ">>" & fieldStr)
                        Else
                            If Not BdfDataBase.DicVddBinIdsFieldMap.Exists(UCase(Replace(CStr(fieldStr), "_shadow", ""))) Then
                                BdfDataBase.DicVddBinIdsFieldMap.Add UCase(Replace(CStr(fieldStr), "_shadow", "")), idsField
                            End If
                        End If
                    End If
                End If
            Next
        Next
#End If
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetVddBinPmodeMap")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function Get_EQN_Voltage_Per_Site(ids_current As Double, p_mode As Long, GRADEVDD As Double, Resolution As Double, PASSBIN As Long, Optional showPrint As Boolean = True, Optional isAddition As Boolean = False) As Long ', ByRef EQN_Number As Long
On Error GoTo errHandler
#If ApplyBinCut Then
Dim passbincut_num As Variant, step As Long
Dim p As Long, i As Long, k As Long
Dim Remainder As Double, Divisor As Double, cal_voltage As Double, final_voltage As Double
Dim m_compareValue As Variant

    Divisor = Resolution
    Get_EQN_Voltage_Per_Site = 999 ''''set to failure
    If BinCut(p_mode, PASSBIN).ExcludedPmode = False Then
    
        If isAddition = True Then
            If (showPrint) Then TheExec.Datalog.WriteComment VddBinName(p_mode) & ", Total Mode_Step=" & (EfuseBinCutAddition(p_mode, PASSBIN).Mode_Step + 1)
            ''''As is : [BinCut(P_mode, PASSBIN).Mode_Step - 1], it will cause that the search is failure.
    
            For k = 0 To EfuseBinCutAddition(p_mode, PASSBIN).Mode_Step
                final_voltage = 0
                If ids_current = 0 Then Exit Function
                cal_voltage = EfuseBinCutAddition(p_mode, PASSBIN).c(k) - EfuseBinCutAddition(p_mode, PASSBIN).m(k) * (log(ids_current * 1000) / log(10))
                Remainder = cal_voltage / Divisor
                Remainder = Floor(Remainder)
                cal_voltage = Remainder * Divisor
    
                final_voltage = cal_voltage + EfuseBinCutAddition(p_mode, PASSBIN).CP_GB(k)
    
                If cal_voltage > EfuseBinCutAddition(p_mode, PASSBIN).CP_Vmax(k) Then
                      final_voltage = EfuseBinCutAddition(p_mode, PASSBIN).CP_Vmax(k) + EfuseBinCutAddition(p_mode, PASSBIN).CP_GB(k)
                ElseIf cal_voltage < EfuseBinCutAddition(p_mode, PASSBIN).CP_Vmin(k) Then
                    final_voltage = EfuseBinCutAddition(p_mode, PASSBIN).CP_Vmin(k) + EfuseBinCutAddition(p_mode, PASSBIN).CP_GB(k)
                End If
    
                If (showPrint) Then
                    TheExec.Datalog.WriteComment "EQ_" & (k + 1) & ", GRADEVDD=" & GRADEVDD & ", IDS_current=" & ids_current & ", Final_voltage=" & final_voltage
                End If
                If GRADEVDD = final_voltage Then
                    Get_EQN_Voltage_Per_Site = k + 1
                    Exit For
                End If
            Next k
        Else
            If (showPrint) Then TheExec.Datalog.WriteComment VddBinName(p_mode) & ", Total Mode_Step=" & (EfuseBinCut(p_mode, PASSBIN).Mode_Step + 1)
            ''''As is : [BinCut(P_mode, PASSBIN).Mode_Step - 1], it will cause that the search is failure.
    
            For k = 0 To EfuseBinCut(p_mode, PASSBIN).Mode_Step
                final_voltage = 0
                If ids_current = 0 Then Exit Function
                cal_voltage = EfuseBinCut(p_mode, PASSBIN).c(k) - EfuseBinCut(p_mode, PASSBIN).m(k) * (log(ids_current * 1000) / log(10))
                Remainder = cal_voltage / Divisor
                Remainder = Floor(Remainder)
                cal_voltage = Remainder * Divisor
    
                final_voltage = cal_voltage + EfuseBinCut(p_mode, PASSBIN).CP_GB(k)
    
                If cal_voltage > EfuseBinCut(p_mode, PASSBIN).CP_Vmax(k) Then
                      final_voltage = EfuseBinCut(p_mode, PASSBIN).CP_Vmax(k) + EfuseBinCut(p_mode, PASSBIN).CP_GB(k)
                ElseIf cal_voltage < EfuseBinCut(p_mode, PASSBIN).CP_Vmin(k) Then
                    final_voltage = EfuseBinCut(p_mode, PASSBIN).CP_Vmin(k) + EfuseBinCut(p_mode, PASSBIN).CP_GB(k)
                End If
    
                If (showPrint) Then
                    TheExec.Datalog.WriteComment "EQ_" & (k + 1) & ", GRADEVDD=" & GRADEVDD & ", IDS_current=" & ids_current & ", Final_voltage=" & final_voltage
                End If
                If GRADEVDD = final_voltage Then
                    Get_EQN_Voltage_Per_Site = k + 1
                    Exit For
                End If
            Next k
        End If
    Else
        TheExec.ErrorLogMessage "The Performance Power for " & VddBinName(p_mode) & " does not exist"
    End If
#End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "Get_EQN_Voltage_Per_Site")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetMapIdsField(powerPin As String) As eFuseBdfField
On Error GoTo errHandler
#If ApplyBinCut Then
Dim bankstr As Variant, bank As eFuseBdfBank
Dim fieldStr As Variant, field As eFuseBdfField
        
    For Each bankstr In BdfDataBase.Banks.Keys
        If bankstr <> Empty Then
            Set bank = BdfDataBase.Banks(bankstr)
            If bank.HadVddBinFuse Then
                For Each fieldStr In bank.Fields.Keys
                    Set field = bank.Fields(fieldStr)
                    If field.Algorithm = alg_ids And LCase(field.name) Like "ids_*_bincheck" Then
                        Set GetMapIdsField = field
                        Exit Function
                    ''202011xxx
                    ElseIf field.Algorithm = alg_ids And _
                            LCase(field.name) Like "ids_" & LCase(powerPin) & "*" Then
                        Set GetMapIdsField = field
                        Exit Function
                    End If
                Next
            End If
        End If
    Next
    Set GetMapIdsField = Nothing
#End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "GetMapIdsField")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub CheckeFusePassBinCutNum()
On Error GoTo errHandler
#If ApplyBinCut Then

    If Not BdfDataBase.BinCutNumBank Is Nothing Then
        If UCase(GlbUtility.currStage) = "CP1" Or UCase(GlbUtility.currStage) = "CP2" Then
            Set CurrentPassBinCutNum = BdfDataBase.BinCutNumBank.GeteFuseValue("product_identifier").Add(1)
        ElseIf UCase(GlbUtility.currStage) <> "WLFT2" Then
            Set CurrentPassBinCutNum = BdfDataBase.BinCutNumBank.GeteFuseValue("product_identifier").Add(1)
        End If
    
        If BdfDataBase.DicVddBinPmodeMap.Count = 0 Then GlbUtility.MessageBox "Didn't contain VddBin-Pmode Mapping!": Exit Sub
    Else
        GlbUtility.MessageBox "Didn't contain " & BinCutIdentifierField & " in any bank! can't obtain CurrentPassBinCutNum"
        Exit Sub
    End If
#End If
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseVddBin_Ids_Glb", "CheckeFusePassBinCutNum")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'202004xx disable for AP end-----

