Attribute VB_Name = "VBT_LIB_DC_AP"
#Const isUFP = True
Option Explicit

Public Function IDS_eFuse_Write(FuseType As String, Flag_Name As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim i As Long
    
    Dim m_len As Long

    Dim m_resolution As Double
    
    Dim showPrint As Boolean: showPrint = True
    
    Dim site As Variant
    Dim MappingPin As Variant
    
    Dim m_valStr As String
    Dim m_dlogstr As String
    Dim m_Fusetype As String
    Dim m_catename As String
    Dim EFUSE_Field_Ary() As String

    Dim m_value As New SiteDouble
    Dim m_decimal As New SiteDouble
    
    Dim Pass_Fail_Flag As New SiteBoolean
    
    Dim m_catename_pinlistdata As New PinListData
    
    Dim EFUSE_IDS_Dic As New Scripting.dictionary
    Dim pin As String
    Dim PinGrp() As String
    Dim Count As Long
    Dim CurrentJob_IDSInfo As IDS_Mapping_Info
    Dim MappingPin_ids As String
    
    For i = 0 To UBound(IDS_Mapping)
        If UCase(currentJobName) = UCase(IDS_Mapping(i).Stage) Then
            CurrentJob_IDSInfo = IDS_Mapping(i)
            Exit For
        End If
    Next i
    
    
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)
    For Each MappingPin In CurrentJob_IDSInfo.MappingDict.Keys   ''MappingPin=vdd_cio
        MappingPin_ids = CurrentJob_IDSInfo.MappingDict.item(MappingPin)    ''MappingPin_ids=ids_vdd_cio_25c
        If MappingPin_ids <> "" Then
''            If Not EFUSE_IDS_Dic.Exists(MappingPin_ids) Then EFUSE_IDS_Dic.item(MappingPin_ids) = gl_IDS_INFO_Dic(MappingPin_ids)       ''Only store from Single Pin, pin group value will from Calc Sum
            If Not EFUSE_IDS_Dic.Exists(LCase(CStr(MappingPin))) Then EFUSE_IDS_Dic.item(LCase(CStr(MappingPin))) = gl_IDS_INFO_Dic(LCase(CStr(MappingPin)))       ''Only store from Single Pin, pin group value will from Calc Sum
            
            m_catename = MappingPin_ids
            
            pin = EFUSE_IDS_Dic.item(LCase(CStr(MappingPin)))(0)
''            theexec.DataManager.DecomposePinList pin, PinGrp(), Count
            
            If Count > 1 Then  ''FT vdd_cio
                m_decimal = GetStoredMeasurement(pin)   ''Value stored from Calc Sum, Pin Group
            Else
''                m_decimal = EFUSE_IDS_Dic.item(CurrentJob_IDSInfo.MappingDict.item(pin))(2)     ''Single Pin value
                m_decimal = EFUSE_IDS_Dic.item(LCase(CStr(MappingPin)))(2)     ''Single Pin value
            End If
            
            Call ids_cal_resolution(FuseType, m_catename, m_decimal, m_value, m_resolution)
                
            For Each site In TheExec.sites
                If TheExec.Flow.SiteFlag(site, Flag_Name) = 1 Then
                    Pass_Fail_Flag(site) = False
                ElseIf TheExec.Flow.SiteFlag(site, Flag_Name) = 0 Then
                    Pass_Fail_Flag(site) = True
                Else
                    Pass_Fail_Flag(site) = False
                    TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
                End If
                  
                If (showPrint) Then
                    m_Fusetype = FuseType
                    m_valStr = Format(m_decimal * 1000, "0.000000")
                    m_Fusetype = FormatNumeric(m_Fusetype, 4)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse IDS_SetWriteDecimal_SetPatTestPass_Flag ", -1)

                    If gB_efuse_DicValue_Chk_Flag = True Then ''Efuse_DicValue_Chk ==>
                        m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_value, -10) + _
                         " (" + FormatNumeric(m_valStr + " mA", 12) + _
                         " / " + Format(m_resolution * 1000, "0.000000") + "mA)" + " @ " + UCase(MappingPin)
                    
                    Else
                        m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_value, -10) + _
                         " (" + FormatNumeric(m_valStr + " mA", 12) + _
                         " / " + Format(m_resolution * 1000, "0.000000") + "mA)"
                    End If ''Efuse_DicValue_Chk <==
                    '------ 'add for Real value validation 211230------------
                   TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            
            '20210406 Modify for new Efuse
            Dim opbank As eFuseBdfBank
            Dim field As eFuseBdfField
            Set opbank = GetBdfBank(FuseType)
            Set field = opbank.Fields(m_catename)
            opbank.SetEfuse field.name, m_decimal, Pass_Fail_Flag, , , , False
        End If
    
    Next MappingPin
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_AP", "IDS_eFuse_Write") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function DCVS_IDS_main_current_Delta(Delta_Pin As PinList, FuseType As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim i As Long, j As Long

    Dim IDS_from_Efuse As New SiteDouble
    Dim IDS_from_DCVS As New SiteDouble
    Dim IDS_Delta As New SiteDouble
    Dim IDS_PwrName As String
    Dim HiLimit_IDS_Delta As Double
    Dim LoLimit_IDS_Delta As Double

    Set gS_delta_IDS_pcpu = New SiteDouble
    Set gS_delta_IDS_ecpu = New SiteDouble
    Set gS_delta_IDS_gpu = New SiteDouble
    Set gS_delta_IDS_dcs_ddr = New SiteDouble
    Set gS_delta_IDS_cpu_sram = New SiteDouble
    Set gS_delta_IDS_ave = New SiteDouble
    
    ''Seperate the Delta_Pin list into Array, Carter - 20190115, Start
    Dim Pins() As String, Pin_Cnt As Long
    Dim CFG_Pins() As String
    TheExec.DataManager.DecomposePinList Delta_Pin, Pins(), Pin_Cnt

    ReDim CFG_Pins(UBound(Pins))
    For i = 0 To UBound(Pins())
        If LCase(Pins(i)) Like "*sram_cpu" Then
            CFG_Pins(i) = "ids_" & LCase(Replace(Pins(i), "SRAM_CPU", "CPU_SRAM"))
        Else
            CFG_Pins(i) = "ids_" & LCase(Pins(i))
        End If
    Next i
    
    If TheExec.TesterMode = testModeOffline Then
        For i = 0 To UBound(Pins())
            IDS_from_DCVS = All_Power_data_IDS_GB.Pins(Pins(i))
            IDS_from_Efuse = IDS_Delta.Add(0.01 + Rnd() * 0.0001)
            IDS_Delta = IDS_from_DCVS.Subtract(IDS_from_Efuse)
            
            TheExec.Flow.TestLimit resultVal:=IDS_from_Efuse, lowVal:=0.0001, Tname:=Pins(i) & "_CP1", PinName:=Pins(i) & "_CP1"
            TheExec.Flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_WLFT", PinName:=Pins(i) & "_WLFT"
            TheExec.Flow.TestLimit IDS_Delta, Tname:=Pins(i) & "_Delta", PinName:=Pins(i) & "_Delta", ForceResults:=tlForceNone
        Next i

    Else
        If FuseType = "CFG" Then
            '20210406 Modify for new Efuse
            Dim opbank As New eFuseBdfBank
            Dim field As New eFuseBdfField
            Dim fieldStr As Variant
            Set opbank = GetBdfBank(FuseType)
            
            For Each fieldStr In opbank.fields  '20210406 Modify for new Efuse
            'For i = 0 To UBound(CFGFuse.category())
                Set field = opbank.fields(fieldStr)
                If field.Algorithm = alg_ids Then   '20210406 Modify for new Efuse
                'If LCase(CFGFuse.category(i).Algorithm) Like "*ids*" Then
                    IDS_PwrName = LCase(field.name) '20210406 Modify for new Efuse
                    'IDS_PwrName = LCase(CFGFuse.category(i).name)
                    ''Do the IDS_Delta if the IDS_PwrName exists in Delta_Pin and Delta_IDS_Dic exists, Carter - 20190116, Start
                    For j = 0 To UBound(Pins())
                        If IDS_PwrName = CFG_Pins(j) Then
                            IDS_from_DCVS = All_Power_data_IDS_GB.Pins(Pins(j))
                            IDS_from_Efuse = field.DsscDecValue.Multiply(field.Resolution * 0.001)  '20210406 Modify for new Efuse
                            'IDS_from_Efuse = CFGFuse.category(i).Read.Decimal.Multiply(CFGFuse.category(i).Resoultion * 0.001)
                            IDS_Delta = IDS_from_DCVS.Subtract(IDS_from_Efuse)
                            If IDS_PwrName Like "*pcpu*" Then
                                gS_delta_IDS_pcpu = IDS_Delta
                            ElseIf IDS_PwrName Like "*ecpu*" Then
                                gS_delta_IDS_ecpu = IDS_Delta
                            ElseIf IDS_PwrName Like "*gpu*" Then
                                gS_delta_IDS_gpu = IDS_Delta
                            ElseIf IDS_PwrName Like "*dcs_ddr*" Then
                                gS_delta_IDS_dcs_ddr = IDS_Delta
                            ElseIf IDS_PwrName Like "*cpu_sram*" Then
                                 gS_delta_IDS_cpu_sram = IDS_Delta
                            ElseIf IDS_PwrName Like "*ave*" Then
                                gS_delta_IDS_ave = IDS_Delta
                            End If
                            TheExec.Flow.TestLimit resultVal:=IDS_from_Efuse, lowVal:=0.0001, Tname:=IDS_PwrName & "_CP1", PinName:=IDS_PwrName & "_CP1"
                            TheExec.Flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_WLFT", PinName:=IDS_PwrName & "_WLFT"
                            TheExec.Flow.TestLimit IDS_Delta, Tname:=IDS_PwrName & "_Delta", PinName:=IDS_PwrName & "_Delta", ForceResults:=tlForceFlow
                            Exit For
                        End If
                    Next j
                    ''Do the IDS_Delta if the IDS_PwrName exists in Delta_Pin and Delta_IDS_Dic exists, Carter - 20190116, End
                End If
            Next
            'Next i
        End If
    End If

'============================================================================================
'=  Record Delta IDS to HardKeyReg (added on 2017/7/10)                                     =
'============================================================================================
    VBT_IEDA_Registry "WLFT_Delta_IDS_PCPU", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_ECPU", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_GPU", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_DCS_DDR", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_CPU_SRAM", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_AVE", True
 
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_AP", "DCVS_IDS_main_current_Delta") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
