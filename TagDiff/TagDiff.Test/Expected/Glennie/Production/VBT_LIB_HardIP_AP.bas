Attribute VB_Name = "VBT_LIB_HardIP_AP"
Option Explicit
''
''Public Function pll_read() As Long
''
''    Call HIP_eFuse_Read(A, b, c)
''    Call HIP_eFuse_Read
''
''End Function


Public Function HIP_eFuse_Write(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, Flag_Name As String, Optional Efuse_Binary_Write_Flag As Boolean = False, _
                                Optional Calc_code As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    ' Parameter : eFuse Block , eFuse Variable , data
    ' Call auto_eFuse_SetPatTestPass_Flag("CFG", "LPDP_C_RX", TheHdw.Digital.Patgen.PatternBurstPassed(Site))
    ' Call auto_eFuse_SetWriteDecimal("CFG", "LPDP_C_RX", BestCode(Site))

    Dim site As Variant
    Dim DSPWave_Dict As New DSPWave
    Dim Data_Temp As String
    Dim m_value As New SiteDouble
    '20220106, Add for Real value validation( reverse bit)
    Dim m_rvalue As New SiteDouble: m_rvalue = 0 ''Efuse_DicValue_Chk ,add for Real value validation 211230
    Dim j As Integer
    Dim Pass_Fail_Flag As New SiteBoolean
    Dim Flag_Name_Split() As String: Flag_Name_Split = Split(Flag_Name, ",")
    Dim Flag_Name_Split_ary_temp() As String
    Dim Flag_Name_Split_ary_sty As String
    Dim i As Long
    Dim fusetype_org As String
    Dim m_dlogstr As String
    Dim opbank As eFuseBdfBank '20210406 Add for new Efuse
    Dim field As eFuseBdfField  '20210406 Add for new Efuse
    On Error GoTo errHandler

    Dim FuseValCheckFail As New SiteBoolean
    FuseValCheckFail = False
    'Added by Oscar for Negative Value check, From JadeCdie, 20201023
    
    Dim k As Long
    Dim lFuseCnt As Long
    
    Dim smcatename_Ary() As String
    Dim mcatename_Tmp As String
    Dim sDictStoreCodeName_Ary() As String
    Dim DictStoreCodeName_Tmp As String
    Dim Efuse_Binary_Write_Flag_ary() As String
    Dim Efuse_Binary_Write_Flag_temp As Boolean
'====================== support mutiple format for this eFuse_Write ==========================
        '''FuseType = CFG
        '''m_catename = pcie_refpll_fcal+lpdp0_auspll_dcocfg+lpdp1_auspll_dcocfg
        '''Dict_Store_Code_Name = pcie_refpll_fcal_ate+lpdp0_auspll_dcocfg_ate+lpdp1_auspll_dcocfg_ate
'=====================================================================================
  
'====================== can't support format as below ================================
        '''FuseType = CFG+CFG+CFG      ==> fix for one bank
        '''m_catename = pcie_refpll_fcal+lpdp0_auspll_dcocfg+lpdp1_auspll_dcocfg
        '''Dict_Store_Code_Name = pcie_refpll_fcal_ate            ==>fuse size is difference with m_catename

'=====================================================================================
    Call Reg_Assign_Processing_eFuse(m_catename, Dict_Store_Code_Name)
                                                
    If InStr(m_catename, "+") <> 0 Then

        smcatename_Ary = Split(m_catename, "+")
        sDictStoreCodeName_Ary = Split(Dict_Store_Code_Name, "+")

        If Not (UBound(smcatename_Ary) = UBound(sDictStoreCodeName_Ary)) Then                     'And UBound(sDictStoreCodeName_Ary) = UBound(Efuse_Binary_Write_Flag_ary)   UBound(Efuse_Binary_Write_Flag_ary) = UBound(Flag_Name_Split) Then
            Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_AP", "HIP_eFuse_Write", "Please check m_catename Argument and Dict_Store_Code_Name,Fuse Size mismatch !!!")
            GoTo errHandler
        End If

        lFuseCnt = UBound(smcatename_Ary)
    Else
        ReDim smcatename_Ary(0)
        ReDim sDictStoreCodeName_Ary(0)
        smcatename_Ary(0) = m_catename
        sDictStoreCodeName_Ary(0) = Dict_Store_Code_Name
        lFuseCnt = 0
    End If
        
    For k = 0 To lFuseCnt


        If UBound(sDictStoreCodeName_Ary) = 0 Then
            DictStoreCodeName_Tmp = sDictStoreCodeName_Ary(0)
            DSPWave_Dict = GetStoredCaptureData(sDictStoreCodeName_Ary(0))
        Else
            DictStoreCodeName_Tmp = sDictStoreCodeName_Ary(k)
            DSPWave_Dict = GetStoredCaptureData(sDictStoreCodeName_Ary(k))
        End If

        If UBound(smcatename_Ary) = 0 Then
            mcatename_Tmp = smcatename_Ary(0)
        Else
            mcatename_Tmp = smcatename_Ary(k)
        End If
        
        
        For Each site In TheExec.sites

            If Efuse_Binary_Write_Flag Then
                For j = 0 To (DSPWave_Dict(site).SampleSize - 1)
                    If DSPWave_Dict(site).Element(j) < 0 Then
                        FuseValCheckFail = True
                        Data_Temp = Data_Temp & 0
                    Else
                        Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(j))
                    End If
                Next j
                m_value(site) = Bin2Dec_rev(Data_Temp)
                m_rvalue(site) = DSPWave_Dict(site).SampleSize ''Efuse_DicValue_Chk,,add for Real value validation 211230''20220106, Add for Real value validation( reverse bit)
                If DSPWave_Dict(site).SampleSize = 1 Then
                    TheExec.Datalog.WriteComment ("***********Error! Bin_DSPWave's sample size is 1*********** ")
                End If
                Data_Temp = vbNullString
            Else
                If DSPWave_Dict(site).Element(0) < 0 Then
                    'Added by Oscar for Negative Value check, From JadeCdie, 20201023
                    DSPWave_Dict(site).Element(0) = 0
                    FuseValCheckFail = True
                Else
                    m_value(site) = DSPWave_Dict(site).Element(0)
                End If
                If DSPWave_Dict(site).SampleSize <> 1 Then
                    TheExec.Datalog.WriteComment ("***********Error! DEC_DSPWave's sample size is <> 1*********** ")
                End If
            End If
        Next site
    '''----------cal write fused code
        If Calc_code <> "" Then
            If Calc_code <> "" Then
            'Calc_code = "add,100"
                If Split(Calc_code, ",")(0) = "add" Then
                    m_value = m_value.Add(Split(Calc_code, ",")(1))
                End If
            End If
        End If
    '''----------cal write fused code
    'Fixed From Sicily, 20200423, Oscar
        If UBound(Flag_Name_Split) > 0 Then
            For Each site In TheExec.sites
                For i = 0 To UBound(Flag_Name_Split_ary_temp)
                    If TheExec.flow.SiteFlag(site, Flag_Name_Split_ary_temp(i)) = -1 Then
                        Pass_Fail_Flag(site) = True
                        TheExec.Datalog.WriteComment ("Error! " & Flag_Name_Split_ary_temp(i) & "(" & site & ")" & " status is Clear !")
                        Exit For
                    ElseIf i = 0 Then
                        Pass_Fail_Flag(site) = TheExec.flow.SiteFlag(site, Flag_Name_Split_ary_temp(i))
                    Else
                        Pass_Fail_Flag(site) = Pass_Fail_Flag(site) Or TheExec.flow.SiteFlag(site, Flag_Name_Split_ary_temp(i))
                    End If
                Next i
                If FuseValCheckFail = True Then
                    'Added by Oscar for Negative Value check, From JadeCdie, 20201023
                    Pass_Fail_Flag(site) = False
                    TheExec.Datalog.WriteComment ("Error! " & DictStoreCodeName_Tmp & " Contains abnormal value!")
                ElseIf Pass_Fail_Flag(site) = True Then
                    Pass_Fail_Flag(site) = False
                Else                               ' Site Boolean Pass_Fail_Flag = False
                    Pass_Fail_Flag(site) = True
                End If
                    
    ''                Else
    ''                    Pass_Fail_Flag(site) = False
    ''                    TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
    '            Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catename, Pass_Fail_Flag(site), True)
    '            Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_value(site), True)
                If (True) Then
                    fusetype_org = ""
                    m_dlogstr = ""
                    fusetype_org = FormatNumeric(FuseType, 4)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + fusetype_org + FormatNumeric("Fuse SetWriteVariable_SiteAware", -35)
                    m_dlogstr = m_dlogstr + FormatNumeric(mcatename_Tmp, Len(fusetype_org)) + " = " + FormatNumeric(m_value, -10)
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
        Else
            For Each site In TheExec.sites
                'Added by Oscar for Negative Value check, From JadeCdie, 20201023
                
                Flag_Name_Split_ary_sty = Flag_Name_Split(0)
                
                If FuseValCheckFail = True Then
                    Pass_Fail_Flag(site) = False
                    TheExec.Datalog.WriteComment ("Error! " & sDictStoreCodeName_Ary(k) & " Contains abnormal value!")
                ElseIf TheExec.flow.SiteFlag(site, Flag_Name_Split_ary_sty) = 1 Then
                    Pass_Fail_Flag(site) = False
                ElseIf TheExec.flow.SiteFlag(site, Flag_Name_Split_ary_sty) = 0 Then
                    Pass_Fail_Flag(site) = True
                Else
                    Pass_Fail_Flag(site) = False
                    TheExec.Datalog.WriteComment ("Error! " & Flag_Name_Split_ary_sty & "(" & site & ")" & " status is Clear !")
                End If
    '            Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catename, Pass_Fail_Flag(site), True)
    '            Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_value(site), True)
                If (True) Then
                    fusetype_org = ""
                    m_dlogstr = ""
                    fusetype_org = FormatNumeric(FuseType, 4)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + fusetype_org + FormatNumeric("Fuse SetWriteVariable_SiteAware", -35)
                    m_dlogstr = m_dlogstr + FormatNumeric(mcatename_Tmp, Len(fusetype_org)) + " = " + FormatNumeric(m_value, -10)
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
        End If
        

    '20210406 Modify for new Efuse
    Set opbank = GetBdfBank(FuseType)
    Set field = opbank.Fields(mcatename_Tmp)
        
    If gB_efuse_DicValue_Chk_Flag = True Then
        opbank.SetEfuse field.name, m_value, Pass_Fail_Flag, , , , True, sDictStoreCodeName_Ary(k), m_rvalue '20220106, Add for Real value validation( reverse bit)
    Else
        opbank.SetEfuse field.name, m_value, Pass_Fail_Flag, , , , True
    End If

    'opbank.SetEfuse field.name, m_value, Pass_Fail_Flag, , , , True, Dict_Store_Code_Name, m_rvalue
'    Call auto_eFuse_SetPatTestPass_Flag_SiteAware(FuseType, m_catename, Pass_Fail_Flag, True)
'    Call auto_eFuse_SetWriteVariable_SiteAware(FuseType, m_catename, m_value, False)
        
    Set m_value = Nothing
    Set DSPWave_Dict = Nothing
           
        
    Next k
        
    ' Check implicit alarms
    TheHdw.Alarms.Check

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Write"
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function TrimUVI80Code_VFI_ADC(Optional pat As String, Optional TestSequence As String, Optional MeasV_PinS As String, Optional MeasI_PinS As String, Optional MeasI_Range As Double, _
    Optional MeasF_PinS_SingleEnd As PinList, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
    Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimFormat As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_DigCapData As String = vbNullString, _
    Optional TrimStoreName As String, Optional Diffaccuracy As Double, Optional TrimMethod As Long, _
    Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional TrimCal_Name As String, Optional Antitrim As Boolean = False, Optional Validating_ As Boolean, Optional Interpose_PrePat As String, Optional CPUA_Flag_In_Pat As Boolean = True, Optional Interpose_PostTest As String, _
    Optional Final_Calc As Boolean = False, Optional TrimPredictStep As Double, Optional NolessthanTarget As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
''    Dim PatCount As Long
    Dim i, j As Integer
    Dim pats() As String
    Dim code() As New SiteLong
    Dim MeasValue As New SiteDouble
    Dim BestCode() As New SiteLong, BestVal() As New SiteDouble, verr() As New SiteDouble, temp As New SiteLong
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim trace As Boolean
    Dim site As Variant
    Dim Ts As Variant
    Dim ADCOUT As New SiteBoolean
    Dim TrimStep As Long
    Dim OutDSP As New DSPWave
    Dim PatCount As Long, PattArray() As String
    Dim TestSequence_array() As String
    Dim doallFlag As Boolean
    Dim finalflag As Boolean
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim TestNameInput As String
    Dim ReCalc As New SiteDouble
    Dim predictFlag As Boolean
    Dim linearFlag As Boolean
    Dim Firstvoltage As New SiteDouble
    Dim lessboolen As New SiteBoolean
    Dim lessboolen_need_change As New SiteBoolean
    
    Dim Temp_sub As Double
    Dim MeasSeqAry() As New SiteDouble
    Dim MeasSeqAry_Best() As New SiteDouble

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Dim TempPat As New Pattern
    TempPat.value = pat
    Call ProcessInputToGLB(TempPat, TestSequence, CPUA_Flag_In_Pat, , , , , MeasV_PinS, MeasF_PinS_SingleEnd.value, MeasF_Interval, MeasF_EventSourceWithTerminationMode, , , , MeasI_PinS, CStr(MeasI_Range), , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_DataWidth) _
                    , CStr(DigSrc_Sample_Size), DigSrc_Equation, DigSrc_Assignment, , , , , CUS_Str_DigCapData, , , , , , , , , , , , , , , Interpose_PrePat, , Interpose_PostTest)
    
    gl_TName_Pat = pat

    Call GetFlowTName

    If Validating_ Then
        Call PrLoadPattern(pat)
        Exit Function    ' Exit after validation
    End If
    Call HardIP_InitialSetupForPatgen
    TestSequence_array = Split(TestSequence, ",")
    
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    PATT_GetPatListFromPatternSet pat, pats, PatCount
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    

    
    Dim SplitByEqual() As String, SplitByColon() As String, TrimCodeSize As Long
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    Dim Trimname() As String
    Dim SplitByComma() As String
    SplitByComma = Split(TrimFormat, ";")
    ReDim Trimname(UBound(SplitByComma))
    ReDim code(UBound(SplitByComma))
    ReDim BestCode(UBound(SplitByComma))
    ReDim BestVal(UBound(SplitByComma))
    ReDim verr(UBound(SplitByComma))
    TrimCodeValue_Min = 0
    For i = 0 To UBound(SplitByComma)
        SplitByEqual = Split(SplitByComma(i), "=")
        SplitByColon = Split(SplitByEqual(1), ":")
        Trimname(i) = SplitByEqual(0)
        TrimCodeSize = SplitByColon(0) + 1
        TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
        If SplitByColon(1) = 0 Then
            TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
        Else
            TrimCodeValue_Max = SplitByColon(1)
        End If
    Next i
    Dim binaryFlag As Boolean
    Dim temp_assignment As String
    temp_assignment = DigSrc_Assignment
     
    '''''''''''''''''''''''''''''''''''Process Trim Method'''''''''''''''''''''''''''''''''''''
    Select Case TrimMethod
        Case 0
            linearFlag = True
            doallFlag = False
        Case 1
            binaryFlag = True
        Case 2
            linearFlag = True
            doallFlag = True
        Case 3
            predictFlag = True
            linearFlag = True
    End Select
    
    
    '''''''''''''''''''''''''''''''''''Binary Search'''''''''''''''''''''''''''''''''''''''''''
    If binaryFlag = True Then
        Dim counter As Long
        Dim trimmax() As New SiteLong
        Dim trimmin() As New SiteLong
        ReDim trimmax(UBound(SplitByComma))
        ReDim trimmin(UBound(SplitByComma))
        For i = 0 To UBound(SplitByComma)
        trimmax(i) = TrimCodeValue_Max
        trimmin(i) = TrimCodeValue_Min
        
        counter = 0
        code(i) = (trimmax(i) + trimmin(i)) / 2
        Next i
        Do While counter < TrimCodeSize
            DigSrc_Assignment = temp_assignment
            
            Call TrimUVI80_Meas_VFI_ADC(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_PinS, MeasValue, MeasI_PinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, MeasSeqAry, Final_Calc)
            For i = 0 To UBound(SplitByComma)
            
            If counter = 0 Then
            verr(i) = MeasSeqAry(i).Subtract(TrimTarget).Abs
            BestCode(i) = code(i)
            BestVal(i) = MeasSeqAry(i)
            End If
            
            For Each site In TheExec.sites
                If MeasSeqAry(i).Subtract(TrimTarget).Abs < verr(i) Then
                    BestCode(i) = code(i)
                    BestVal(i) = MeasSeqAry(i)
                    verr(i) = MeasSeqAry(i).Subtract(TrimTarget).Abs
'                    For j = 0 To UBound(MeasSeqAry)
'                        MeasSeqAry_Best(j) = MeasSeqAry(j)
'                    Next j
                    'TheExec.DataLog.WriteComment "Site : " & Site & ",Best Value : " & BestVal(i)
                    'TheExec.DataLog.WriteComment "Site : " & Site & ",Best Code : " & BestCode(i)
                End If
            Next site
            
            For Each site In TheExec.sites
                If Antitrim = True Then
                    If MeasSeqAry(i)(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(i)(site) = code(i)(site) - 1
                        Else
                            trimmax(i)(site) = code(i)(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(i)(site) = code(i)(site)
                        Else
                            trimmin(i)(site) = code(i)(site)
                        End If
                    End If
                Else
                    If MeasSeqAry(i)(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(i)(site) = code(i)(site)
                        Else
                            trimmin(i)(site) = code(i)(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(i)(site) = code(i)(site) - 1
                        Else
                            trimmax(i)(site) = code(i)(site)
                        End If
                    End If
                End If
            Next site
            
            
            If counter = TrimCodeSize - 1 Then
            Else
                code(i) = trimmax(i).Add(trimmin(i)).divide(2)
            End If
            
            
            Next i
            counter = counter + 1
        Loop
        
        DigSrc_Assignment = temp_assignment
        Call TrimUVI80_Meas_VFI_ADC(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_PinS, MeasValue, MeasI_PinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, MeasSeqAry, Final_Calc)
        For i = 0 To UBound(SplitByComma)
            For Each site In TheExec.sites
                    If MeasSeqAry(i).Subtract(TrimTarget).Abs < verr(i) Then
                        BestCode(i) = code(i)
                        BestVal(i) = MeasSeqAry(i)
                        verr(i) = MeasSeqAry(i).Subtract(TrimTarget).Abs
                    End If
                Next site
        Next i
            
    Else
    'Do nothing
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
    
    
    finalflag = True
    
    'Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, BestCode, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodesize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, finalflag)
    
    
EndTrim:

    If Interpose_PostTest <> "" Then
        Call SetForceCondition(Interpose_PostTest & ";STOREPREPAT")
    End If
    
    For i = 0 To UBound(SplitByComma)
    If MeasV_PinS <> "" Then

        TestNameInput = Report_TName_From_Instance("V", vbNullString, "TrimmedVoltage", 0, CLng(i), "3=" & CStr(i))

        If InStr(glb_TestInstance, "T4P2") <> 0 Then
            ReCalc = BestVal(i).Add(1).Multiply(0.7975).Add(0.4)
            'OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            TestNameInput = Merge_TName(OutputTname_format)
            TheExec.flow.TestLimit ReCalc, , , , , , unitVolt, , TestNameInput, , MeasV_PinS, , , , , tlForceFlow
            'OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            TestNameInput = Merge_TName(OutputTname_format)
            TheExec.flow.TestLimit BestVal(i), , , , , , unitVolt, , TestNameInput, , , , , , , tlForceFlow
        Else
            TheExec.flow.TestLimit BestVal(i), , , , , , unitVolt, , TestNameInput, , MeasV_PinS, , , , , tlForceFlow
        End If
    ElseIf MeasI_PinS <> "" Then

        TestNameInput = Report_TName_From_Instance("I", vbNullString, "TrimmedVoltage", 0, CLng(i), "3=" & CStr(i))

        TheExec.flow.TestLimit BestVal(i), , , , , , unitAmp, , TestNameInput, , MeasI_PinS, , , , , tlForceFlow
    ElseIf MeasF_PinS_SingleEnd <> "" Then

        TestNameInput = Report_TName_From_Instance("F", vbNullString, "TrimmedFrequency", 0, CLng(i), "3=" & CStr(i))
        TheExec.flow.TestLimit BestVal(i), , , , , , unitHz, , TestNameInput, , MeasF_PinS_SingleEnd, , , , , tlForceFlow
    Else

        TestNameInput = Report_TName_From_Instance("C", vbNullString, "TrimmedCode(Decimal)", 0, CLng(i), "3=" & CStr(i))

        TheExec.flow.TestLimit BestVal, , , , , , unitNone, , TestNameInput, , DigCap_Pin, , , , , tlForceFlow
    End If

        TestNameInput = Report_TName_From_Instance("C", vbNullString, "TrimmedCode", 0, CLng(i), "3=" & CStr(i))

    TheExec.flow.TestLimit BestCode(i), 0, 2 ^ TrimCodeSize, , , , , , TestNameInput, , , , , , , tlForceNone 'Un-Used_
    'ByPassTestLimit = False
    Next i
    If DigCap_Sample_Size <> 0 Then
        Dim DigCapPinAry() As String, NumberPins As Long
        
        'Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, BestCode, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodesize, Trimname, Meas_StoreName, Cal_Eqn, TrimCal_Name)
        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
        
        If NumberPins > 1 Then
            'Call CreateSimulateDataDSPWave_Parallel(OutDSP, DigCap_Sample_Size)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP, NumberPins)
            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, NumberPins)

        ElseIf NumberPins = 1 Then
            'Call CreateSimulateDataDSPWave(OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP, NumberPins)
            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
        Else 'Do nothing '20230601
        End If
    End If
    

    Dim TempVal As Integer
    Dim FinalTrimCode() As New DSPWave
    Dim SplitTrimStoreName() As String
    SplitTrimStoreName = Split(TrimStoreName, "+")
    
    'FinalTrimCode.CreateConstant 0, TrimCodeSize
    ReDim FinalTrimCode(UBound(SplitTrimStoreName)) As New DSPWave
    For i = 0 To UBound(SplitTrimStoreName)
    FinalTrimCode(i).CreateConstant 0, TrimCodeSize
    For Each site In TheExec.sites
        TempVal = BestCode(i)(site)
        For j = 0 To TrimCodeSize - 1
            FinalTrimCode(i).Element(j) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next j
    Next site
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(SplitTrimStoreName(i), FinalTrimCode(i))
    End If
    
    Next i
    If TrimCal_Name <> "" Then
        If Final_Calc = True Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    Else
        If Calc_Eqn <> "" Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    End If
    
    
    Call HardIP_WriteFuncResult(, , LCase(glb_TestInstance))
    
    DebugPrintFunc pat
    

    
Exit Function 'Add ErrHandler 2023/05/29
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_AP", "TrimUVI80Code_VFI_ADC") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TrimCodeFreq(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasF_PinS_SingleEnd As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double = 1000000, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PreMeas As String, _
    Optional Interpose_PrePat As String, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_MainProgram As String = vbNullString, _
    Optional CUS_Str_DigCapData As String = vbNullString, Optional CUS_Str_DigSrcData As String = vbNullString, Optional MSB_First_Flag As Boolean = False, _
    Optional Trimming_Direction_Increase As Boolean, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    'If Trimming_Direction_Increase = True => measure data < target will add 1 code      '20231018 update
    'Update from T-BraC and T-Har -- 20230628
    Dim PatCount As Long, PattArray() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPwave As New DSPWave
    
    Call ProcessInputToGLB(patset, TestSequence, True, , , , , , MeasF_PinS_SingleEnd, , , , , , , , , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, DigSrc_Assignment, , , , CUS_Str_MainProgram, , , , , , , , , , , , , , , , Interpose_PrePat, Interpose_PreMeas)
    
    'If Trimming_Direction_Increase = "" Then Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP_AP", "TrimCodeFreq", "Trimming_Direction_Increase empty, Please check and add the direction!!") ' Add Warning message -- 20231018
    
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Current Trimming Direction is: " & Trimming_Direction_Increase & "****************")
    
    Dim site As Variant
    Dim pat As String
    Dim i As Long, j As Long, k As Long, p As Long
    
    Dim MeasureFreq As New PinListData, MeasureFreq_F1 As New PinListData, MeasureFreq_F2 As New PinListData
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    
    
    If MSB_First_Flag = True Then
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Note: DSSC MSB First. Please check pattern content"
    End If
    Call GetFlowTName
         
    ''Update Interpose_PreMeas 20170801
    Dim Interpose_PreMeas_Ary() As String
    ''20160923 - Analyze Interpose_PreMeas to force setting with different sequence.
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    Dim l_DigSrc_Sample_Size As Long
    If InStr(DigSrc_Sample_Size, "|") <> 0 Then
        l_DigSrc_Sample_Size = Split(DigSrc_Sample_Size, "|")(1)
    Else
        l_DigSrc_Sample_Size = CLng(DigSrc_Sample_Size)
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    If (UBound(TestSequenceArray) > UBound(TName_Ary)) Then
        ReDim Preserve TName_Ary(UBound(TestSequenceArray)) As String
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    TheHdw.patterns(patset).Load
    gl_TName_Pat = patset.value

''    Call PATT_GetPatListFromPatternSet(patSet.value, PattArray, PatCount)
    
    Dim tempPatArr() As String
    Dim tempPatCnt As Long
    Dim tempSrcSigName As String
    Dim tempSrcPatSeq As Long
    Dim tempVarArray As Variant
    tempSrcPatSeq = 1

    Call PatternBurstCheckAndSplit(patset.value, PattArray, PatCount)
    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        Call PATT_GetPatListFromPatternSet(patset.value, tempPatArr, tempPatCnt)
        tempSrcSigName = Split(tempPatArr(tempSrcPatSeq), ":")(1)
    End If
    
    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        tempVarArray = TheHdw.DSSC.pins(DigSrc_pin).Pattern(tempPatArr(1)).Source.Labels.list  ''20210609 temp
    Else
        tempVarArray = TheHdw.DSSC.pins(DigSrc_pin).Pattern(PattArray(0)).Source.Labels.list  ''20210609 temp
    End If
    
    tempSrcSigName = tempVarArray(0)
    If tempSrcSigName = "" Then
        tempSrcSigName = "TrimCodeCap"
    End If
   
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim d_MeasF_Interval  As Double
    d_MeasF_Interval = 0.001
    
    Dim b_HighThanTargetFreq As New SiteBoolean
    b_HighThanTargetFreq = False
    
    Dim OutputTrimCode As String
    Dim TestLimitIndex As Long, LastSectionF1F2_Index As Long
    LastSectionF1F2_Index = 0
    
    ''==================================================================================================
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    '' 20160706 Create value for final frequency
    Dim b_DefineFinalFreq As New SiteBoolean
    Dim FinalFreq As New PinListData
    
    ''20160712 - If match taget freq just store the trim code
    Dim b_MatchTagetFreq As New SiteBoolean
    Dim b_DisplayFreq As New SiteBoolean
    Dim StoredTargetTrimCode As New DSPWave
    b_MatchTagetFreq = False
    b_DisplayFreq = False
''    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim StoreEachTrimFreq() As New PinListData
    Dim StoreEachTrimCode() As New DSPWave
''    ReDim StoreEachTrimFreq(DigSrc_Sample_Size + 1) As New PinListData
''    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False

    ''20170721-Updated the TrimStart when the first bit is zero and seperate with "&"
    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", vbNullString)
    End If

    Dec_TrimStart_1st = Bin2Dec(TrimStart)

    InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong
    
    Dim InDspWave_Temp As New DSPWave
    Dim TrimFormat_RegSize As Long
    Dim TrimFormat_StoreName As String
    Dim TrimFormat_Ary() As String
    If DigSrc_Equation <> "" Or DigSrc_Assignment <> "" Then
        If InStr(TrimFormat, "=") <> 0 Then
            TrimFormat_Ary = Split(TrimFormat, "=")
            TrimFormat_StoreName = TrimFormat_Ary(0)
            TrimFormat_RegSize = Split(TrimFormat_StoreName, "_")(1)
            Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, TrimFormat_RegSize, InDspWave_Temp)
            Call StoreDataAllType(TrimFormat_StoreName, InDspWave_Temp)
        End If
    End If
    Dim Rtn_SweepTestName As String
    
    
    StoredTargetTrimCode.CreateConstant 0, TrimFormat_RegSize, DspLong
    ReDim StoreEachTrimFreq(TrimFormat_RegSize + 1) As New PinListData
    ReDim StoreEachTrimCode(TrimFormat_RegSize + 1) As New DSPWave
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, TrimFormat_RegSize, DspLong
    Next i
    
    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        Call GeneralDigSrcSettingWithBurst(LCase(patset.value), DigSrc_pin, InDSPwave, Rtn_SweepTestName, MSB_First_Flag)
    Else
        Call TrimCodeFreq_GeneralDigSrcSetting(PattArray(0), DigSrc_pin, l_DigSrc_Sample_Size, l_DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, , , MSB_First_Flag)
    End If
    
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("First Time Setup")
    '' Debug use
    For Each site In TheExec.sites
        OutputTrimCode = vbNullString
        For k = 0 To InDSPwave(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
    Next site
    
    For Each site In TheExec.sites
        StoreEachTrimCode(0)(site) = InDspWave_Temp(site).COPY
    Next site
    
     ''20211125 start process DigCap
    Dim OutDspWave As New DSPWave
    Set OutDspWave = Nothing
    If DigCap_Pin <> "" Then
        Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, 224, OutDspWave)
        ''20211125 end
    End If
    Call TheHdw.patterns(PattArray(0)).start
    
    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0
    Dim temp_array_MeasF_PinS_SingleEnd() As String
    Dim temp_MeasF_PinS_SingleEnd As New PinList
    temp_array_MeasF_PinS_SingleEnd = Split(MeasF_PinS_SingleEnd, "+")
    
    For Each Ts In TestSequenceArray
        temp_MeasF_PinS_SingleEnd = temp_array_MeasF_PinS_SingleEnd(TestSeqNum)
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20160923 - Add Interpose_PreMeas entry point by each sequence
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            Else 'Do nothing '20230601
            End If
        End If
        
        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(temp_MeasF_PinS_SingleEnd, d_MeasF_Interval)
            TheHdw.Digital.pins(temp_MeasF_PinS_SingleEnd).Levels.DriverMode = tlDriverModeVt
            Call HardIP_Freq_MeasFreqStart(temp_MeasF_PinS_SingleEnd, d_MeasF_Interval, MeasureFreq, 0.001)
            
            If TheExec.TesterMode = testModeOffline Then
                Call SimulateOutputFreq(temp_MeasF_PinS_SingleEnd, MeasureFreq)
            End If
        ElseIf UCase(Ts) = "N" Then ''''''//*****Modify For Grackle. William 20210423 *****//
            'do nothing
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20161206-Restore force condiction after measurement
        ''Call SetForceCondition("RESTORE")
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            Else 'Do nothing '20230601
            End If
        End If

        TestSeqNum = TestSeqNum + 1

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    TheHdw.Digital.Patgen.HaltWait
    
    StoreEachTrimFreq(0) = MeasureFreq
    
    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
    PastDSPWave = InDspWave_Temp
    
    TestNameInput = "Freq_meas_"
    TestLimitIndex = 0
    
    '' 20160712 - Modify to use WriteComment to display output frequency.
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.pins(0).value(site) / 1000000), 6) & "M Hz")
        Next site
    End If
    '' 20160712 - Compare Measure Frequency whether match target Freq
    b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
    
    b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
    For Each site In TheExec.sites
        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            StoredTargetTrimCode(site) = InDspWave_Temp(site).COPY
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
    
    
    ''========================================================================================
    ''20161128 Pre check Min/Max trim code process
    Dim b_KeepGoing As New SiteBoolean
    Dim PreviousFreq As New PinListData
    If PreCheckMinMaxTrimCode = True Then
        PreviousFreq = MeasureFreq
        Call rundsp.PreCheckMinMaxTrimCode(b_HighThanTargetFreq, InDspWave_Temp)
''        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPWave)
'        Call TrimCodeFreq_GeneralDigSrcSetting(tempPatArr(1), DigSrc_pin, l_DigSrc_Sample_Size, l_DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPWave, , , MSB_First_Flag)
        If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
            Call GeneralDigSrcSettingWithBurst(LCase(patset.value), DigSrc_pin, InDSPwave, Rtn_SweepTestName, MSB_First_Flag)
        Else
            Call TrimCodeFreq_GeneralDigSrcSetting(PattArray(0), DigSrc_pin, l_DigSrc_Sample_Size, l_DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, , , MSB_First_Flag)
        End If
        
        ''Update Interpose_PreMeas 20170801
        TestSeqNum = 0
        
        For Each Ts In TestSequenceArray
            temp_MeasF_PinS_SingleEnd = temp_array_MeasF_PinS_SingleEnd(TestSeqNum)
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
            Else
                Call TheHdw.Digital.Patgen.HaltWait
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                Else 'Do nothing '20230601
                End If
            End If
            
            If UCase(Ts) = "F" Then
                Call Freq_MeasFreqSetup(temp_MeasF_PinS_SingleEnd, d_MeasF_Interval)
                Call HardIP_Freq_MeasFreqStart(temp_MeasF_PinS_SingleEnd, d_MeasF_Interval, MeasureFreq, 0.001)
                
                If TheExec.TesterMode = testModeOffline Then
                    Call SimulatePreCheckOutputFreq(temp_MeasF_PinS_SingleEnd, MeasureFreq)
                End If
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20161206-Restore force condiction after measurement
            ''Call SetForceCondition("RESTORE")
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                Else 'Do nothing '20230601
                End If
            End If
    
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA)
            Else
                TheHdw.Digital.Patgen.HaltWait
            End If
        Next Ts
        
        TheHdw.Digital.Patgen.HaltWait
        
        For Each site In TheExec.sites
            OutputTrimCode = vbNullString
            For k = 0 To InDSPwave(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
            Next k
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
        Next site
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.pins(0).value(site) / 1000000), 6) & "M Hz")
            Next site
        End If
        For Each site In TheExec.sites
            If b_HighThanTargetFreq(site) = True Then
                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(LessThan, 0)
            Else
                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(GreaterThan, 0)
            End If
        Next site

        Dim PreCheckBinStr As String, PreCheckDecVal As Double
        For Each site In TheExec.sites
            If b_KeepGoing(site) = False Then
                b_StopTrimCodeProcess(site) = True
                PreCheckBinStr = vbNullString
                StoredTargetTrimCode(site) = InDspWave_Temp(site).COPY
                For i = 0 To StoredTargetTrimCode(site).SampleSize - 1
                    PreCheckBinStr = PreCheckBinStr & StoredTargetTrimCode.Element(i)
                Next i
                PreCheckDecVal = Bin2Dec_rev_Double(PreCheckBinStr)
                ''TheExec.Flow.TestLimit PreCheckDecVal, 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TheExec.DataManager.InstanceName & "_TrimCode_Decimal", ForceResults:=tlForceFlow
            End If
        Next site
    Else
        'Do nothing
    End If
    
    ''========================================================================================
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 1
    
    ''20170103-Setup b_KeepGoing to true if PreCheckMinMaxTrimCode=false
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If
    
    If b_KeepGoing.All(False) Then
    Else

        For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)
            
            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                SrcStartBit = SrcStartBit + 1
            End If
            
            For j = SrcStartBit To SrcEndBit Step -1
            
                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                Else
                ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
    ''                b_ControlNextBit = False
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If
    
                If b_FirstExecution = True And j = SrcEndBit Then
                    If Trimming_Direction_Increase = True Then
                        Call rundsp.SetupTrimCodeBit(PastDSPWave, False, j, b_ControlNextBit, InDspWave_Temp)
                    ElseIf Trimming_Direction_Increase = False Then
                        Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDspWave_Temp)
                    End If
                Else
                    If Trimming_Direction_Increase = False Then
                        Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq.LogicalNot, j, b_ControlNextBit, InDspWave_Temp)
                    ElseIf Trimming_Direction_Increase = True Then
                        Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDspWave_Temp)
                    End If
                End If
                
                If DigSrc_Equation <> "" Or DigSrc_Assignment <> "" Then
                    If InStr(TrimFormat, "=") <> 0 Then
                        Call StoreDataAllType(TrimFormat_StoreName, InDspWave_Temp)
                    End If
                End If

''                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPWave)
'                Call TrimCodeFreq_GeneralDigSrcSetting(tempPatArr(1), DigSrc_pin, l_DigSrc_Sample_Size, l_DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPWave, , , MSB_First_Flag)
                If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
                    Call GeneralDigSrcSettingWithBurst(LCase(patset.value), DigSrc_pin, InDSPwave, Rtn_SweepTestName, MSB_First_Flag)
                Else
                    Call TrimCodeFreq_GeneralDigSrcSetting(PattArray(0), DigSrc_pin, l_DigSrc_Sample_Size, l_DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, , , MSB_First_Flag)
                End If
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site).data = InDspWave_Temp(site).data
                Next site
            
                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule
                
                If gl_Disable_HIP_debug_log = False Then
                    If b_FirstExecution = True Then
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    Else
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j)
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    End If
                End If
                
                For Each site In TheExec.sites
    ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                    If b_KeepGoing(site) = True Then
                        OutputTrimCode = vbNullString
                        For k = 0 To InDSPwave(site).SampleSize - 1
                            OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
                        Next k
                        
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
                    End If
    ''                End If
                Next site
                '' ==============================================================================================
                
                Call TheHdw.patterns(PattArray(0)).start
                
                ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0
                
                For Each Ts In TestSequenceArray
                    temp_MeasF_PinS_SingleEnd = temp_array_MeasF_PinS_SingleEnd(TestSeqNum)
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
                    Else
                        Call TheHdw.Digital.Patgen.HaltWait
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20160923 - Add Interpose_PreMeas entry point by each sequence
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                        Else 'Do nothing '20230601
                        End If
                    End If
                    
                    
                    If UCase(Ts) = "F" Then
                        Call Freq_MeasFreqSetup(temp_MeasF_PinS_SingleEnd, d_MeasF_Interval)
                        Call HardIP_Freq_MeasFreqStart(temp_MeasF_PinS_SingleEnd, d_MeasF_Interval, MeasureFreq, 0.001)
                        
                        '--------------- off line mode data --------
                        If TheExec.TesterMode = testModeOffline Then
                            Dim SimuIndex As Long
                            SimuIndex = TestLimitIndex
                            If SimuIndex >= 8 Then
                                SimuIndex = 8
                            End If
                            Call SimulateOutputFreq(temp_MeasF_PinS_SingleEnd, MeasureFreq)
                            MeasureFreq.pins(temp_MeasF_PinS_SingleEnd).value(0) = MeasureFreq.pins(temp_MeasF_PinS_SingleEnd).value(0) - (SimuIndex * 1000)
                        End If
                        '--------------------------------------------
                        
                        If j = SrcEndBit + 1 Then
                            MeasureFreq_F1 = MeasureFreq
                        ElseIf j = SrcEndBit Then
                            MeasureFreq_F2 = MeasureFreq
                        Else 'Do nothing '20230601
                        End If
                    Else
                        '' Do nothing
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20161206-Restore force condiction after measurement
                    ''Call SetForceCondition("RESTORE")
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        Else 'Do nothing '20230601
                        End If
                    End If
            
                    TestSeqNum = TestSeqNum + 1
                    
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                    Else
                        TheHdw.Digital.Patgen.HaltWait
                    End If
                Next Ts
                
                TheHdw.Digital.Patgen.HaltWait
                
                StoreEachTrimFreq(StoreEachIndex) = MeasureFreq
                StoreEachIndex = StoreEachIndex + 1
                
                If j = SrcEndBit Then
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq_F1.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureFreq_F2.Math.Subtract(TrimTarget).Abs)
                    If Trimming_Direction_Increase = False Then
                        Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq.LogicalNot, j, b_ControlNextBit, InDspWave_Temp)
                    ElseIf Trimming_Direction_Increase = True Then
                        Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDspWave_Temp)
                    End If
                    
                    PastDSPWave = InDspWave_Temp
                Else
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                    PastDSPWave = InDspWave_Temp
                End If
    
                TestLimitIndex = TestLimitIndex + 1
                
                '' 20160712 - Modify to use WriteComment to display output frequency.
                
                If gl_Disable_HIP_debug_log = False Then
                
                    For Each site In TheExec.sites
        ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        If b_KeepGoing(site) = True Then
                            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.pins(0).value(site) / 1000000), 6) & "M Hz")
                        End If
        ''                End If
                    Next site
                End If
                
                ''20160716 - Modify display info sequence when source bit in the section end
                If j = SrcEndBit Then
                    For Each site In TheExec.sites
    ''                    If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        
                        If b_KeepGoing(site) = True And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 1 & " Output Frequency = " & FormatNumber((MeasureFreq_F1.pins(0).value(site) / 1000000), 6) & "M Hz")
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 2 & " Output Frequency = " & FormatNumber((MeasureFreq_F2.pins(0).value(site) / 1000000), 6) & "M Hz")
                        End If
    ''                    End If
                    Next site
                    LastSectionF1F2_Index = LastSectionF1F2_Index + 2
                End If
                
                '' 20160712 - Compare Measure Frequency whether match target Freq
                b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                            StoredTargetTrimCode(site) = InDspWave_Temp(site).COPY
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                Next site
                ''20161128-Stop trim code process if found out match code of all site
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
            Next j
        Next i
    End If
    
    ''============================================================================
    ''20161128 Findout mimiumn trim code
    Dim CloseTargetFreq As New PinListData
    Dim DiffValue As New SiteLong, PreviousDiffValue As New SiteLong, CloseIndex As New SiteLong
    
    Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteLong
    Dim binstr As String
    Dim CloseTargetTrimCode As New DSPWave
    Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
    Dim b_FirstTimeSwitch As Boolean
    
    If b_KeepGoing.All(False) Then
    Else
''        If TrimPrcocessAll = True Then
            CloseTargetTrimCode.CreateConstant 0, TrimFormat_RegSize, DspLong
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If StoredTargetTrimCode(site).CalcSum = 0 Then
                        b_UseMinTrim(site) = True
                    End If
                End If
            Next site
            
            If UseMinimumTrimCode = True Then
                b_UseMinTrim = True
            End If
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_UseMinTrim(site) = True Then
                        '' Findout minimum difference value
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            DiffValue(site) = Abs(StoreEachTrimFreq(i).pins(0).value(site) - TrimTarget)
                            If DiffValue(site) <= PreviousDiffValue(site) Then
                                CloseIndex(site) = i
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                            If i = 0 Then
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                        Next i
                        '' Transfer to decimal value to findout minimum code
                        PreviousDecVal = 0
                        DecVal = 0
                        b_FirstTimeSwitch = False
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            binstr = vbNullString
                            If Abs(StoreEachTrimFreq(i).pins(0).value(site) - TrimTarget) = MinDiffVal(site) Then
                                For j = 0 To StoreEachTrimCode(i)(site).SampleSize - 1
                                    binstr = binstr & StoreEachTrimCode(i)(site).Element(j)
                                Next j
                                DecVal = Bin2Dec_rev_Double(binstr)
                               
                                If DecVal < PreviousDecVal Then
                                    MinDecVal = DecVal
                                    CloseTargetTrimCode(site).data = StoreEachTrimCode(i)(site).data
                                End If
                                PreviousDecVal = DecVal
                                If b_FirstTimeSwitch = False Then
                                    CloseTargetTrimCode(site).data = StoreEachTrimCode(i)(site).data
                                    b_FirstTimeSwitch = True
                                End If
                            End If
                        Next i
                    Else
                        'Do nothing
                    End If
                End If
            Next site
''        End If
    End If
    
    For Each site In TheExec.sites
        If b_KeepGoing(site) = True Then
            If b_UseMinTrim(site) = True Then
                StoredTargetTrimCode(site) = CloseTargetTrimCode(site).COPY
            Else
                StoredTargetTrimCode(site) = StoredTargetTrimCode(site).COPY
            End If
        Else
            StoredTargetTrimCode(site) = StoredTargetTrimCode(site).COPY
        End If
    Next site
    ''============================================================================
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, StoredTargetTrimCode)
    End If
    
    
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)

    For Each site In TheExec.sites
        OutputTrimCode = vbNullString
        For k = 0 To StoredTargetTrimCode(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(StoredTargetTrimCode(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Final Output Trim Code = " & OutputTrimCode)
    Next site
    
    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel StoredTargetTrimCode, TrimFormat_RegSize, ConvertedDataWf
    
    If DigSrc_Equation <> "" Or DigSrc_Assignment <> "" Then
        If InStr(TrimFormat, "=") <> 0 Then
            Call StoreDataAllType(TrimFormat_StoreName, StoredTargetTrimCode)
        End If
    End If
    
    TestNameInput = Report_TName_From_Instance(CalcC, DigSrc_pin.value, vbNullString, 0, 0)
    
    TheExec.flow.TestLimit ConvertedDataWf.Element(0), 0, 2 ^ TrimFormat_RegSize - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
'    Call TrimCodeFreq_GeneralDigSrcSetting(tempPatArr(1), DigSrc_pin, l_DigSrc_Sample_Size, l_DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPWave, , , MSB_First_Flag)
    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        Call GeneralDigSrcSettingWithBurst(LCase(patset.value), DigSrc_pin, InDSPwave, Rtn_SweepTestName, MSB_First_Flag)
    Else
        Call TrimCodeFreq_GeneralDigSrcSetting(PattArray(0), DigSrc_pin, l_DigSrc_Sample_Size, l_DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, , , MSB_First_Flag)
    End If
    Call TheHdw.patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
    
    For Each Ts In TestSequenceArray
        temp_MeasF_PinS_SingleEnd = temp_array_MeasF_PinS_SingleEnd(TestSeqNum)
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20160923 - Add Interpose_PreMeas entry point by each sequence
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            Else 'Do nothing '20230601
            End If
        End If
        
        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(temp_MeasF_PinS_SingleEnd, d_MeasF_Interval)
            Call HardIP_Freq_MeasFreqStart(temp_MeasF_PinS_SingleEnd, d_MeasF_Interval, MeasureFreq, 0.001)
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20161206-Restore force condiction after measurement
        ''Call SetForceCondition("RESTORE")
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            Else 'Do nothing '20230601
            End If
        End If

        TestSeqNum = TestSeqNum + 1
        
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    
    TheHdw.Digital.Patgen.HaltWait
    
    If TPModeAsCharz_GLB Then
        Disable_Inst_pinname_in_PTR
        TheExec.flow.TestLimit resultVal:=MeasureFreq, unit:=unitHz, ForceResults:=tlForceFlow
        Enable_Inst_pinname_in_PTR
    Else
        For p = 0 To MeasureFreq.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("F", MeasureFreq.pins(p), "Final", CInt(p))
            TheExec.flow.TestLimit resultVal:=MeasureFreq, unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next p
    End If
    
    Dim sl_FUSE_Val As New SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else

    End If
    
    DebugPrintFunc patset.value
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
    

    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_AP", "TrimCodeFreq") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function





Public Function HIP_eFuse_Read(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, dspwavesize As Long, Optional Efuse_Read_Dec_Flag As Boolean = False, Optional Dict_Store_Dec_Name As String = vbNullString, _
                                Optional Calc_code As String = vbNullString, Optional NonZero_Val_Chk As String = vbNullString) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    ' Parameter : eFuse Block , eFuse Variable , data , Data Width
    ' Create dictionary , if exist then remove and re-create
    ' MUST :  if necessary , we can set limit if read out value = 0 then bin out .

    Dim site As Variant
    Dim Read_Code As New DSPWave
    Dim Read_Value As New DSPWave
    Dim Efuse_Value As New SiteLong
    Dim TempVal As Long
    Dim Efuse_Value_Chk As New SiteVariant
    Dim i As Long
    
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim m_catename_Ary() As String
    Dim Dict_Store_Code_Name_Ary() As String
    Dim dspwavesize_Ary() As String
    Dim k As Long
    Dim ReadCnt As Long
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Call Reg_Assign_Processing_eFuse(m_catename, Dict_Store_Code_Name)
    
    If InStr(m_catename, "+") <> 0 Then
        
        m_catename_Ary = Split(m_catename, "+")
        Dict_Store_Code_Name_Ary = Split(Dict_Store_Code_Name, "+")


        If Not (UBound(m_catename_Ary) = UBound(m_catename_Ary)) Then
            Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_AP", "HIP_eFuse_Read", "Please check m_catename Argument size, Fuse Read times mismatch !")
            GoTo errHandler
        End If
        ReadCnt = UBound(m_catename_Ary)
    Else
        
        ReDim m_catename_Ary(0)
        ReDim Dict_Store_Code_Name_Ary(0)
        

        m_catename_Ary(0) = m_catename
        Dict_Store_Code_Name_Ary(0) = Dict_Store_Code_Name
        
        ReadCnt = 0
    End If
    
    For k = 0 To ReadCnt
    
        Set Read_Code = Nothing
                
        Read_Code.CreateConstant 0, dspwavesize
    
        If UCase(Efuse_Read_Dec_Flag) = "TRUE" Then
            Read_Value.CreateConstant 0, 1
        End If
        
    '    TheExec.Datalog.WriteComment "error: bypass HIP efuse read, wait for efuse check!"
    '    Exit Function
        '20210406 Modify for new Efuse
        'If LCase(m_catename) Like "mon*" Then FuseType = "MON"
        Efuse_Value = GetEfuseHipValue(FuseType, m_catename_Ary(k))
        
        Dim Val_from_Store As New DSPWave
        Dim Data_Temp As String
    
        For Each site In TheExec.sites
    
            If UCase(currentJobName) = "CP1" Then
                If TheExec.enableWord("MTR_value_read_back") And Efuse_Value(site) = 0 Then
                    Val_from_Store = GetStoredCaptureData(Dict_Store_Code_Name_Ary(k))
                    Data_Temp = vbNullString
                    For i = 0 To (Val_from_Store(site).SampleSize - 1)
                        If Val_from_Store(site).Element(i) < 0 Then
                            Data_Temp = Data_Temp & 0
                        Else
                            Data_Temp = Data_Temp & (Val_from_Store(site).Element(i))
                        End If
                    Next i
                    Efuse_Value(site) = Bin2Dec_rev(Data_Temp)
                    TheExec.Datalog.WriteComment ("site" & site & " (" & Dict_Store_Code_Name_Ary(k) & ")" & " Read fuse = 0, use Trim value = " & Efuse_Value(site))
                End If
            End If
    
    
            'Efuse_Value(site) = auto_eFuse_GetReadDecimal(FuseType, m_catename, True)
    '''''        Efuse_Value(Site) = CLng(Site) + 8
    '''----------cal get fused code
            If Calc_code <> "" Then
            'Calc_code = "minus,100"
                If Split(Calc_code, ",")(0) = "minus" Then
                    Efuse_Value = Efuse_Value.Subtract(Split(Calc_code, ",")(1))
                End If
            End If
    '''----------cal get fused code
            If UCase(Efuse_Read_Dec_Flag) = "TRUE" Then
                Read_Value.Element(0) = Efuse_Value(site)
            End If
    
            TempVal = Efuse_Value(site)
            For i = 0 To dspwavesize - 1
                Read_Code.Element(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
    
            If UCase(glb_TestInstance) Like "MTRGR_VREF*" Then
                Dim Read_Code_Inverse As New DSPWave: Read_Code_Inverse = Read_Code
                If Read_Code.Element(dspwavesize - 1) = 0 Then
                    Read_Code_Inverse.Element(dspwavesize - 1) = 1
                Else
                    Read_Code_Inverse.Element(dspwavesize - 1) = 0
                End If
            End If
    
            If Efuse_Value(site) = 0 Then
            'If Read out value = 0 then bin out
                Efuse_Value_Chk(site) = 0
            Else
                Efuse_Value_Chk(site) = 1
            End If
    
        Next site
        'NonZero_Val_Chk = "false"
        'Replace tlforceFlow to tlforceNone, refer donan - 20230621
        'Update Trim Code allow to be zero case, extend spec avoid result show Fail -- 20220819
        If LCase(NonZero_Val_Chk) = "true" Then
            TheExec.flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone
        Else
            TheExec.Datalog.WriteComment "Trim Code allow to be zero, set NonZero_Val_Chk low limit:=0"
            TheExec.flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=0, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone
        End If
    
    
    'Update Trim Code allow to be zero case, extend spec avoid result show Fail -- 20220819
    '    If LCase(NonZero_Val_Chk) <> "false" Then
    '        TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone
    '    ElseIf LCase(TheExec.DataManager.instanceName) = "rlxosc_fuse_read_rlxosc_auto_fcal_ft1" And UCase(TheExec.CurrentJob) Like "WLFT1" Then
    '        TheExec.Datalog.WriteComment "Error: Tahiti WLFT1 pipe clean need to read fuse"
    '        TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=0, hiVal:=0, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone
    '    Else
    '        TheExec.Datalog.WriteComment "Trim Code allow to be zero, set NonZero_Val_Chk low limit:=0"
    '        TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=0, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone
    '    End If
        
        'Added by Oscar for Standard read name requested by HDC, From JadeCdie, 20201023
        If gl_UseStandardTestName_Flag = True Then
            'Call Report_TName_From_Instance(Calc, "", OutputTname_format, "X", "", "", 0, 0)
            Call Report_ALG_TName_From_Instance(OutputTname_format, Calc, "X", vbNullString, 0)
                OutputTname_format(1) = "FuseRead"
                OutputTname_format(2) = "x"
                OutputTname_format(3) = FuseType & "-" & OutputTname_format(4) ' fuse type
                OutputTname_format(4) = "TRIM"
                OutputTname_format(5) = Replace(m_catename_Ary(k), "_", "-") ' fuse name
                OutputTname_format(6) = Replace(Dict_Store_Code_Name_Ary(k), "_", "-") ' dictionary name
                OutputTname_format(9) = "x"
                
            TestNameInput = Merge_TName(OutputTname_format)
            TestNameInput = left(TestNameInput, Len(TestNameInput) - 1)
        
            TheExec.flow.TestLimit Efuse_Value, , , , , , , , Tname:=TestNameInput, ForceResults:=tlForceNone
        End If
        'Added by Oscar, From JadeCdie, 20201023
        
        Call AddStoredCaptureData(Dict_Store_Code_Name_Ary(k), Read_Code)
        If UCase(glb_TestInstance) Like "MTRGR_VREF*" Then: Call AddStoredCaptureData("var_bg_rref", Read_Code_Inverse)
        If UCase(Efuse_Read_Dec_Flag) = "TRUE" Then
            Call AddStoredCaptureData(Dict_Store_Dec_Name, Read_Value)
        End If
    Next k
    Set Read_Code = Nothing
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Read"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function TrimCodeFreq_New_ALG(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double = 1000000, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PreMeas As String, _
    Optional TrimCodeRepeat As Integer, _
    Optional Interpose_PrePat As String, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function
    End If


    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)

    Dim PatCount As Long, PattArray() As String

    Call GetFlowTName
    Call HardIP_InitialSetupForPatgen
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.patterns(patset).Load

    'Call PATT_GetPatListFromPatternSet(PatSet.value, PattArray, PatCount)
    Dim tempPatArr() As String
    Dim tempPatCnt As Long
    Dim tempSrcSigName As String
    Dim tempSrcPatSeq As Long
    tempSrcPatSeq = 1

    Call PatternBurstCheckAndSplit(patset.value, PattArray, PatCount)
    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        Call PATT_GetPatListFromPatternSet(patset.value, tempPatArr, tempPatCnt)
        tempSrcSigName = Split(tempPatArr(tempSrcPatSeq), ":")(1)
    End If

    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TrimCodeRepeat <= 0 Then TrimCodeRepeat = 1

    ''
    Dim l_DigSrc_Sample_Size As Long
    If InStr(DigSrc_Sample_Size, "|") <> 0 Then
        l_DigSrc_Sample_Size = Split(DigSrc_Sample_Size, "|")(1)
    Else
        l_DigSrc_Sample_Size = CLng(DigSrc_Sample_Size)
    End If

    Dim i As Long, j As Long, k As Long, p As Long
    Dim b_MatchTagetFreq As New SiteBoolean, b_DisplayFreq As New SiteBoolean
    Dim StoredTargetTrimCode As New DSPWave, StoreEachTrimCode() As New DSPWave
    Dim StoreEachTrimFreq() As New PinListData
    ReDim StoreEachTrimFreq(l_DigSrc_Sample_Size + 1) As New PinListData
    ReDim StoreEachTrimCode(l_DigSrc_Sample_Size + 1) As New DSPWave
    b_MatchTagetFreq = False
    b_DisplayFreq = False
    StoredTargetTrimCode.CreateConstant 0, l_DigSrc_Sample_Size, DspLong
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, l_DigSrc_Sample_Size, DspLong
    Next i
    Dim StoreEachIndex As Long
    Dim PastDSPWave As New DSPWave, InDSPwave As New DSPWave
    Dim ExecutionMax As Long, SrcStartBit As Long, SrcEndBit As Long
    Dim d_MeasF_Interval  As Double
    d_MeasF_Interval = 0.001 ''20190903 0.001
    'd_MeasF_Interval = 0.01 ''20190903 0.001
    Dim b_HighThanTargetFreq As New SiteBoolean
    b_HighThanTargetFreq = False
    Dim TestLimitIndex As Long, LastSectionF1F2_Index As Long
    LastSectionF1F2_Index = 0
    Dim b_DefineFinalFreq As New SiteBoolean
    Dim FinalFreq As New PinListData
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    Dim Interpose_PreMeas_Ary() As String, TestSequenceArray() As String
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    TestSequenceArray = Split(TestSequence, ",")
    Dim MeasureFreq_F1 As New PinListData, MeasureFreq_F2 As New PinListData
    Dim Ts As Variant
    Dim TestNameInput As String
    Dim b_KeepGoing As New SiteBoolean
    Dim PreviousFreq As New PinListData
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    Dim CloseTargetFreq As New PinListData
    Dim DiffValue As New SiteLong, PreviousDiffValue As New SiteLong, CloseIndex As New SiteLong

    Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteLong
    Dim binstr As String
    Dim CloseTargetTrimCode As New DSPWave
    Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
    Dim b_FirstTimeSwitch As Boolean
    ''

    Dim site As Variant
    Dim MeasureFreq As New PinListData
    Dim OutputTrimCode As String
    Dim InitialDSPWave As New DSPWave, TrimStartDSPWave As New DSPWave, InDSPwave_Repeat As New DSPWave
    Dim StrSeparatebyComma() As String, StrSeparatebyEqual() As String, StrSeparatebyColon() As String

    Dim StartBit_CoarseTuning As Long, EndBit_CoarseTuning As Long
    Dim TrimStart_Dec As Long
    Dim TempCnt As Long
    Dim MaskDSPWave0 As New DSPWave, MaskDSPWave1 As New DSPWave
    Dim bTrimCodeFreqContinue As New SiteBoolean, bCoarseTuningContinue As New SiteBoolean, bCheckAverageContinue As New SiteBoolean
    Dim InitialDSPWave_F0 As New DSPWave, InitialDSPWave_F1 As New DSPWave
    Dim StoreEachTrimFreq0() As New PinListData, StoreEachTrimFreq1() As New PinListData
    Dim StoreTrimFreq0 As New PinListData, StoreTrimFreq1 As New PinListData
    ReDim StoreEachTrimFreq0(l_DigSrc_Sample_Size)
    ReDim StoreEachTrimFreq1(l_DigSrc_Sample_Size)
    Dim bF0HigherThanTargetFreq As New SiteBoolean, bF1HigherThanTargetFreq As New SiteBoolean
    Dim MeasureFreq0 As New PinListData, MeasureFreq1 As New PinListData
    Dim FinalDSPCoarseTuning As New DSPWave
    FinalDSPCoarseTuning.CreateConstant 0, l_DigSrc_Sample_Size, DspLong
    Dim FreqAVG01 As New PinListData, FreqAVG23 As New PinListData
    Dim bF2HigherThanTargetFreq As New SiteBoolean, bF3HigherThanTargetFreq As New SiteBoolean
    Dim bFAVG01HigherThanTargetFreq As New SiteBoolean, bFAVG23HigherThanTargetFreq As New SiteBoolean
    Dim bFAVG01EqualToTargetFreq As New SiteBoolean
    Dim bTempSiteSelected As New SiteBoolean


    Dim tempVarArray As Variant
    Dim sSrcSigName As String


    StrSeparatebyComma = Split(TrimFormat, ";")
    StrSeparatebyEqual = Split(StrSeparatebyComma(0), "=")
    StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
    StartBit_CoarseTuning = StrSeparatebyColon(0)
    EndBit_CoarseTuning = StrSeparatebyColon(1)

    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", vbNullString)
    End If
    TrimStart_Dec = Bin2Dec(TrimStart)
    TrimStartDSPWave.CreateConstant TrimStart_Dec, 1, DspLong
    Call rundsp.CreateFlexibleDSPWave(TrimStartDSPWave, l_DigSrc_Sample_Size, InitialDSPWave)
    For Each site In TheExec.sites
        InitialDSPWave(site) = InitialDSPWave(site).ConvertDataTypeTo(DspLong)
    Next site
    TheExec.Datalog.WriteComment ("Initial Trim Code")
    Call TrimCodeFreq_WriteComment_DspTrimCode(InitialDSPWave)

    'Coarse Tuning
    TheExec.Datalog.WriteComment ("********** Coarse Tuning")
    bCoarseTuningContinue = True
    bTrimCodeFreqContinue = True

    MaskDSPWave0.CreateConstant 0, l_DigSrc_Sample_Size, DspLong
    MaskDSPWave1.CreateConstant 1, l_DigSrc_Sample_Size, DspLong
    For TempCnt = StartBit_CoarseTuning To EndBit_CoarseTuning Step -1
        MaskDSPWave0.Element(TempCnt) = 1 '1111 0000 0000
        MaskDSPWave1.Element(TempCnt) = 0 '0000 1111 1111
    Next TempCnt

    Dim TempCnt_TrimStep As Long
    For TempCnt_TrimStep = StartBit_CoarseTuning To EndBit_CoarseTuning Step -1
        TheExec.Datalog.WriteComment ("======================================================================================")
        TheExec.Datalog.WriteComment ("BIT" & TempCnt_TrimStep & "=1")
        For Each site In TheExec.sites
            InitialDSPWave_F0(site) = InitialDSPWave(site).bitwiseand(MaskDSPWave0(site))
            InDSPwave_Repeat(site) = InitialDSPWave_F0(site).Repeat(TrimCodeRepeat)
        Next site
        Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
''        Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
''            Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, tempSrcSigName, DigSrc_Sample_Size, InDSPwave)
            Call TrimCodeFreq_RunPat_and_MeasF(tempPatArr(tempSrcPatSeq), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        Else
            Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        End If

        StoreEachTrimFreq0(TempCnt_TrimStep) = MeasureFreq
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency0 = " & FormatNumber((StoreEachTrimFreq0(TempCnt_TrimStep).pins(0).value(site) / 1000000), 6) & "M Hz")
        Next site

        For Each site In TheExec.sites
            InitialDSPWave_F1(site) = InitialDSPWave(site).BitwiseOr(MaskDSPWave1(site))
            InDSPwave_Repeat(site) = InitialDSPWave_F1(site).Repeat(TrimCodeRepeat)
        Next site
        Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
''        Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
            Call TrimCodeFreq_RunPat_and_MeasF(tempPatArr(tempSrcPatSeq), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        Else
            Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        End If

        StoreEachTrimFreq1(TempCnt_TrimStep) = MeasureFreq
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency1 = " & FormatNumber((StoreEachTrimFreq1(TempCnt_TrimStep).pins(0).value(site) / 1000000), 6) & "M Hz")
        Next site

        bF0HigherThanTargetFreq = StoreEachTrimFreq0(TempCnt_TrimStep).Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        bF1HigherThanTargetFreq = StoreEachTrimFreq1(TempCnt_TrimStep).Math.Subtract(TrimTarget).compare(GreaterThan, 0)

        For Each site In TheExec.sites
            If (bF0HigherThanTargetFreq(site) = False) And (bF1HigherThanTargetFreq(site) = False) Then
                If TempCnt_TrimStep = EndBit_CoarseTuning Then
                    TheExec.Datalog.WriteComment ("Site " & site & ": EXIT FUNCTION !!!")
                    ''20190903
                    StoredTargetTrimCode(site) = InitialDSPWave(site)
                    bCoarseTuningContinue(site) = False
                    bTrimCodeFreqContinue(site) = False
                End If
                InitialDSPWave(site).Element(TempCnt_TrimStep - 1) = 1
            ElseIf bF0HigherThanTargetFreq(site) = True Then
                If TempCnt_TrimStep = EndBit_CoarseTuning Then
                    TheExec.Datalog.WriteComment ("Site " & site & ": EXIT FUNCTION !!!")
                    ''20190903
                    StoredTargetTrimCode(site) = InitialDSPWave(site)
                    bCoarseTuningContinue(site) = False
                    bTrimCodeFreqContinue(site) = False
                End If
                InitialDSPWave(site).Element(TempCnt_TrimStep) = 0
                InitialDSPWave(site).Element(TempCnt_TrimStep - 1) = 1
            ElseIf (bF0HigherThanTargetFreq(site) = False) And (bF1HigherThanTargetFreq(site) = True) Then
                bCoarseTuningContinue(site) = False
                StoreTrimFreq0 = StoreEachTrimFreq0(TempCnt_TrimStep)
                StoreTrimFreq1 = StoreEachTrimFreq1(TempCnt_TrimStep)
                TheExec.Datalog.WriteComment ("Site " & site & ": FO<Target AND F1>Target, Go Check Average")
            Else 'Do nothing '20230601
            End If
        Next site
        If bCoarseTuningContinue.All(False) = True Then
            Exit For
        End If
        TheExec.sites.Selected = bCoarseTuningContinue
    Next TempCnt_TrimStep

    'Check Average
    TheExec.Datalog.WriteComment ("********** Check Average")
    ''TEST
    ''bTrimCodeFreqContinue(5) = False
    'TheExec.sites.Selected = True
    TheExec.sites.Selected = bTrimCodeFreqContinue
    bCheckAverageContinue = bTrimCodeFreqContinue
    FinalDSPCoarseTuning = InitialDSPWave

    bTempSiteSelected = TheExec.sites.Selected
    If bTempSiteSelected.Any(True) = True Then
        TheExec.Datalog.WriteComment ("Recode F_Coarse")
        FreqAVG01 = StoreTrimFreq0.Math.Add(StoreTrimFreq1).divide(2)
        bFAVG01HigherThanTargetFreq = FreqAVG01.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        bFAVG01EqualToTargetFreq = FreqAVG01.Math.Subtract(TrimTarget).compare(EqualTo, 0) ''20190903
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " AVG01(F0,F1) = " & FormatNumber((FreqAVG01.pins(0).value(site) / 1000000), 6) & "M Hz")
        Next site
        Dim TempDSP As New DSPWave
        TempDSP.CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
            TempDSP(site) = InitialDSPWave(site).ConvertStreamTo(tldspParallel, l_DigSrc_Sample_Size, 0, Bit0IsMsb)
            ''
            If TempDSP(site).Element(0) < 2 ^ (EndBit_CoarseTuning) Or TempDSP(site).Element(0) > (2 ^ (StartBit_CoarseTuning + 1) - 1 - 2 ^ (EndBit_CoarseTuning)) Then ''256 to 2 ^ (EndBit_CoarseTuning + 1), 3839
                TheExec.Datalog.WriteComment ("Site " & site & " : ERROR")
            End If
            ''
            If bFAVG01HigherThanTargetFreq = True Then
                TempDSP(site) = TempDSP(site).Add(-1 * 2 ^ (EndBit_CoarseTuning)).ConvertDataTypeTo(DspLong) 'watch out 0000XXXXXXXX
                InitialDSPWave(site) = TempDSP(site).ConvertStreamTo(tldspSerial, l_DigSrc_Sample_Size, 0, Bit0IsMsb)
                InitialDSPWave(site) = InitialDSPWave(site).ConvertDataTypeTo(DspLong)
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01>Target, CurrentTrimCode-1")
            ElseIf bFAVG01HigherThanTargetFreq = False And bFAVG01EqualToTargetFreq = False Then ''20190903
                TempDSP(site) = TempDSP(site).Add(2 ^ (EndBit_CoarseTuning)).ConvertDataTypeTo(DspLong) 'watch out 1111XXXXXXXX
                InitialDSPWave(site) = TempDSP(site).ConvertStreamTo(tldspSerial, l_DigSrc_Sample_Size, 0, Bit0IsMsb)
                InitialDSPWave(site) = InitialDSPWave(site).ConvertDataTypeTo(DspLong)
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01<Target, CurrentTrimCode+1")
            ElseIf bFAVG01EqualToTargetFreq = True Then ''20190903
                bCheckAverageContinue(site) = False
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01=Target, Go Fine Tuning")
            Else 'Do nothing '20230601
            End If
        Next site
    End If

    'TheExec.sites.Selected = True
    TheExec.sites.Selected = bTrimCodeFreqContinue.LogicalAnd(bCheckAverageContinue)
    TheExec.Datalog.WriteComment ("********** Check Average : F2 F3")
    bTempSiteSelected = TheExec.sites.Selected
    If bTempSiteSelected.Any(True) = True Then
        For Each site In TheExec.sites
            InitialDSPWave_F0(site) = InitialDSPWave(site).bitwiseand(MaskDSPWave0(site))
            InDSPwave_Repeat(site) = InitialDSPWave_F0(site).Repeat(TrimCodeRepeat)
        Next site
        Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)


        'Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        'Fixed by Rhodes 20210628 CCPANJ
        If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
             Call TrimCodeFreq_RunPat_and_MeasF(tempPatArr(tempSrcPatSeq), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        Else
             Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        End If






        StoreEachTrimFreq0(TempCnt_TrimStep) = MeasureFreq
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency2 = " & FormatNumber((StoreEachTrimFreq0(TempCnt_TrimStep).pins(0).value(site) / 1000000), 6) & "M Hz")
        Next site

        For Each site In TheExec.sites
            InitialDSPWave_F1(site) = InitialDSPWave(site).BitwiseOr(MaskDSPWave1(site))
            InDSPwave_Repeat(site) = InitialDSPWave_F1(site).Repeat(TrimCodeRepeat)
        Next site
        Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
''        Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
            Call TrimCodeFreq_RunPat_and_MeasF(tempPatArr(tempSrcPatSeq), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        Else
            Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        End If

        StoreEachTrimFreq1(TempCnt_TrimStep) = MeasureFreq
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency3 = " & FormatNumber((StoreEachTrimFreq1(TempCnt_TrimStep).pins(0).value(site) / 1000000), 6) & "M Hz")
        Next site

        bF0HigherThanTargetFreq = StoreEachTrimFreq0(TempCnt_TrimStep).Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        bF1HigherThanTargetFreq = StoreEachTrimFreq1(TempCnt_TrimStep).Math.Subtract(TrimTarget).compare(GreaterThan, 0)

        For Each site In TheExec.sites
            If bF0HigherThanTargetFreq(site) = True Or bF1HigherThanTargetFreq(site) = False Then
                bCheckAverageContinue(site) = False
                TheExec.Datalog.WriteComment ("Site " & site & " : F2>Target OR F3<Target, Go Fine Tuning")
            End If
        Next site
    End If

    TheExec.Datalog.WriteComment ("********** Check Average : FreqAVG01 vs FreqAVG23")
    'TheExec.sites.Selected = True
    TheExec.sites.Selected = bTrimCodeFreqContinue.LogicalAnd(bCheckAverageContinue)
    bTempSiteSelected = TheExec.sites.Selected
    If bTempSiteSelected.Any(True) = True Then
        FreqAVG23 = StoreEachTrimFreq0(TempCnt_TrimStep).Math.Add(StoreEachTrimFreq1(TempCnt_TrimStep)).divide(2)
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " AVG23(F2,F3) = " & FormatNumber((FreqAVG23.pins(0).value(site) / 1000000), 6) & "M Hz")
        Next site
        FreqAVG01 = FreqAVG01.Math.Subtract(TrimTarget).Abs
        FreqAVG23 = FreqAVG23.Math.Subtract(TrimTarget).Abs
        For Each site In TheExec.sites
            If FreqAVG01.pins(0).value(site) > FreqAVG23.pins(0).value(site) Then ''20190903upadte, from "<" to ">"
                FinalDSPCoarseTuning(site) = InitialDSPWave(site)
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01>FreqAVG23,         Update F_Coarse, Go Fine Tuning")
            Else
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01<=FreqAVG23, Do not Update F_Coarse, Go Fine Tuning")
            End If
        Next site
    End If

    'Go Fine Tuning

    'Fine Tuning
    TheExec.Datalog.WriteComment ("********** Fine Tuning")
    TheExec.sites.Selected = True
    If bTrimCodeFreqContinue.All(False) = True Then
        TheExec.Datalog.WriteComment ("********** All Site Fail !!!, Jump to Final")
        GoTo TrimCodeTTR
    Else
        TheExec.sites.Selected = bTrimCodeFreqContinue
    End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    TrimFormat = StrSeparatebyComma(1)
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    For Each site In TheExec.sites
        InDSPwave(site) = FinalDSPCoarseTuning(site)
        InDSPwave_Repeat(site) = InDSPwave(site).Repeat(TrimCodeRepeat)
    Next site
    Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
    ''

    'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)

    '/* ------ Added by Kaino on 2019/06/12 ------- */
    For Each site In TheExec.sites
        InDSPwave_Repeat = InDSPwave.Repeat(TrimCodeRepeat)
    Next site
''    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
'        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'    Else
'        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'    End If
    '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */


    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("First Time Setup")
    '' Debug use
    For Each site In TheExec.sites
        OutputTrimCode = vbNullString

        If gl_Disable_HIP_debug_log = False Then
            'For k = 0 To InDSPwave(Site).SampleSize - 1
            '    OutputTrimCode = OutputTrimCode & CStr(InDSPwave(Site).Element(k))
            'Next k

            '/* ------ Added by Kaino on 2019/06/12 ------- */
            For k = 0 To InDSPwave_Repeat(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave_Repeat(site).Element(k))
               If (k + 1) Mod l_DigSrc_Sample_Size = 0 And k < InDSPwave_Repeat(site).SampleSize - 1 Then
                   OutputTrimCode = OutputTrimCode & ","
               End If
            Next k
        '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */

            TheExec.Datalog.WriteComment ("Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
                End If
    Next site

    For Each site In TheExec.sites
        StoreEachTrimCode(0)(site) = InDSPwave(site).COPY
    Next site

    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then

        tempVarArray = TheHdw.DSSC.pins(DigSrc_pin).Pattern(tempPatArr(tempSrcPatSeq)).Source.Labels.list ''20210609 temp
        sSrcSigName = tempVarArray(0)
'        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, sSrcSigName, l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
        Call TheHdw.patterns(tempPatArr(tempSrcPatSeq)).start
    Else
        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
    Call TheHdw.patterns(PattArray(0)).start
    End If

    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0

    For Each Ts In TestSequenceArray
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If

        ''Update Interpose_PreMeas 20170801
        ''20160923 - Add Interpose_PreMeas entry point by each sequence
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            Else 'Do nothing '20230601
            End If
        End If

        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)

            If TheExec.TesterMode = testModeOffline Then
                Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
            End If
        End If

        ''Update Interpose_PreMeas 20170801
        ''20161206-Restore force condiction after measurement
        ''Call SetForceCondition("RESTORE")
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            Else 'Do nothing '20230601
            End If
        End If

        TestSeqNum = TestSeqNum + 1

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    TheHdw.Digital.Patgen.HaltWait

    StoreEachTrimFreq(0) = MeasureFreq

    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
    PastDSPWave = InDSPwave

    TestNameInput = "Freq_meas_"
    TestLimitIndex = 0

    '' 20160712 - Modify to use WriteComment to display output frequency.
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.pins(0).value(site) / 1000000), 6) & "M Hz")
        Next site
    End If
    '' 20160712 - Compare Measure Frequency whether match target Freq
    b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)

    b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
    For Each site In TheExec.sites
        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            StoredTargetTrimCode(site) = InDSPwave(site).COPY
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")


    ''========================================================================================
    ''20161128 Pre check Min/Max trim code process
'    Dim b_KeepGoing As New SiteBoolean
'    Dim PreviousFreq As New PinListData
    If PreCheckMinMaxTrimCode = True Then
        PreviousFreq = MeasureFreq
        Call rundsp.PreCheckMinMaxTrimCode(b_HighThanTargetFreq, InDSPwave)


        'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
        '/* ------ Added by Kaino on 2019/06/12 ------- */
        For Each site In TheExec.sites
            InDSPwave_Repeat = InDSPwave.Repeat(TrimCodeRepeat)
        Next site
''        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'        If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
'            Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'        Else
'            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'        End If
        '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */

        ''Update Interpose_PreMeas 20170801
    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        tempVarArray = TheHdw.DSSC.pins(DigSrc_pin).Pattern(tempPatArr(tempSrcPatSeq)).Source.Labels.list ''20210609 temp
        sSrcSigName = tempVarArray(0)
'        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, sSrcSigName, l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
        Call TheHdw.patterns(tempPatArr(tempSrcPatSeq)).start
    Else
        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
        Call TheHdw.patterns(PattArray(0)).start
    End If

        TestSeqNum = 0

        For Each Ts In TestSequenceArray
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
            Else
                Call TheHdw.Digital.Patgen.HaltWait
            End If

            ''Update Interpose_PreMeas 20170801
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                Else 'Do nothing '20230601
                End If
            End If

            If UCase(Ts) = "F" Then
                Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)

                If TheExec.TesterMode = testModeOffline Then
                    Call SimulatePreCheckOutputFreq(MeasureF_Pin, MeasureFreq)
                End If
            End If

            ''Update Interpose_PreMeas 20170801
            ''20161206-Restore force condiction after measurement
            ''Call SetForceCondition("RESTORE")
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                Else 'Do nothing '20230601
                End If
            End If

            TestSeqNum = TestSeqNum + 1

            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA)
            Else
                TheHdw.Digital.Patgen.HaltWait
            End If
        Next Ts

        TheHdw.Digital.Patgen.HaltWait

        For Each site In TheExec.sites
            OutputTrimCode = vbNullString
            For k = 0 To InDSPwave(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
            Next k
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
        Next site

        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.pins(0).value(site) / 1000000), 6) & "M Hz")
            Next site
        End If
        For Each site In TheExec.sites
            If b_HighThanTargetFreq(site) = True Then
                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(LessThan, 0)
            Else
                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(GreaterThan, 0)
            End If
        Next site

        Dim PreCheckBinStr As String, PreCheckDecVal As Double
        For Each site In TheExec.sites
            If b_KeepGoing(site) = False Then
                b_StopTrimCodeProcess(site) = True
                PreCheckBinStr = vbNullString
                StoredTargetTrimCode(site) = InDSPwave(site).COPY
                For i = 0 To StoredTargetTrimCode(site).SampleSize - 1
                    PreCheckBinStr = PreCheckBinStr & StoredTargetTrimCode.Element(i)
                Next i
                PreCheckDecVal = Bin2Dec_rev_Double(PreCheckBinStr)
                ''TheExec.Flow.TestLimit PreCheckDecVal, 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TheExec.DataManager.InstanceName & "_TrimCode_Decimal", ForceResults:=tlForceFlow
            End If
        Next site
    End If

    ''========================================================================================




'    Dim b_ControlNextBit As Boolean
'    b_ControlNextBit = False
'    Dim b_FirstExecution As Boolean
'    b_FirstExecution = False
    StoreEachIndex = 1

    ''20170103-Setup b_KeepGoing to true if PreCheckMinMaxTrimCode=false
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If

    If b_KeepGoing.All(False) Then
    Else

        For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)

            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                SrcStartBit = SrcStartBit + 1
            End If

            For j = SrcStartBit To SrcEndBit Step -1

                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                Else
                ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
    ''                b_ControlNextBit = False
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If

                If b_FirstExecution = True And j = SrcEndBit Then
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPwave)
    ''            ElseIf b_FirstExecution = False And j = SrcStartBit Then
    ''                j = SrcStartBit + 1
                Else
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                End If


                'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
                '/* ------ Added by Kaino on 2019/06/12 ------- */
                For Each site In TheExec.sites
                    InDSPwave_Repeat = InDSPwave.Repeat(TrimCodeRepeat)
                Next site
''                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'                If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
'                    Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'                Else
'                    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'                End If
                '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */

                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site) = InDSPwave(site).COPY
                Next site

                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule

                If gl_Disable_HIP_debug_log = False Then
                    If b_FirstExecution = True Then
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    Else
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j)
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    End If
                End If

                For Each site In TheExec.sites
    ''              If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                    If b_KeepGoing(site) = True Then
                                                If gl_Disable_HIP_debug_log = False Then
                                OutputTrimCode = vbNullString
                                'For k = 0 To InDSPwave(Site).SampleSize - 1
                                '    OutputTrimCode = OutputTrimCode & CStr(InDSPwave(Site).Element(k))
                                'Next k
                                '/* ------ Added by Kaino on 2019/06/12 ------- */
                                For k = 0 To InDSPwave_Repeat(site).SampleSize - 1
                                    OutputTrimCode = OutputTrimCode & CStr(InDSPwave_Repeat(site).Element(k))
                                    If (k + 1) Mod l_DigSrc_Sample_Size = 0 And k < InDSPwave_Repeat(site).SampleSize - 1 Then
                                        OutputTrimCode = OutputTrimCode & ","
                                    End If
                                Next k
                                '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */
                                TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
                                                End If
                    End If
    ''              End If
                Next site
                '' ==============================================================================================

                    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
                        tempVarArray = TheHdw.DSSC.pins(DigSrc_pin).Pattern(tempPatArr(tempSrcPatSeq)).Source.Labels.list ''20210609 temp
                        sSrcSigName = tempVarArray(0)
                '        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
                        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, sSrcSigName, l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
                        Call TheHdw.patterns(tempPatArr(tempSrcPatSeq)).start
                    Else
                        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
                Call TheHdw.patterns(PattArray(0)).start
                    End If

                ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0

                For Each Ts In TestSequenceArray
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
                    Else
                        Call TheHdw.Digital.Patgen.HaltWait
                    End If

                    ''Update Interpose_PreMeas 20170801
                    ''20160923 - Add Interpose_PreMeas entry point by each sequence
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                        Else 'Do nothing '20230601
                        End If
                    End If


                    If UCase(Ts) = "F" Then
                        Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                        Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)

                        '--------------- off line mode data --------
                        If TheExec.TesterMode = testModeOffline Then
                            Dim SimuIndex As Long
                            SimuIndex = TestLimitIndex
                            If SimuIndex >= 8 Then
                                SimuIndex = 8
                            End If
                            Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
                            MeasureFreq.pins(MeasureF_Pin).value(0) = MeasureFreq.pins(MeasureF_Pin).value(0) - (SimuIndex * 1000)
                           ' MeasureFreq.Pins(MeasureF_Pin).Value(1) = MeasureFreq.Pins(MeasureF_Pin).Value(1) + (SimuIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(2) = MeasureFreq.Pins(MeasureF_Pin).Value(2) + (TestLimitIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(3) = MeasureFreq.Pins(MeasureF_Pin).Value(3) - (TestLimitIndex * 1000)
                        End If
                        '--------------------------------------------

                        If j = SrcEndBit + 1 Then
                            MeasureFreq_F1 = MeasureFreq
                        ElseIf j = SrcEndBit Then
                            MeasureFreq_F2 = MeasureFreq
                        Else 'Do nothing '20230601
                        End If
                    Else
                        '' Do nothing
                    End If

                    ''Update Interpose_PreMeas 20170801
                    ''20161206-Restore force condiction after measurement
                    ''Call SetForceCondition("RESTORE")
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        Else 'Do nothing '20230601
                        End If
                    End If

                    TestSeqNum = TestSeqNum + 1

                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                    Else
                        TheHdw.Digital.Patgen.HaltWait
                    End If
                Next Ts

                TheHdw.Digital.Patgen.HaltWait



                Call TrimCodeFreq_RunPat_and_MeasF(tempPatArr(tempSrcPatSeq), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, l_DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)




                StoreEachTrimFreq(StoreEachIndex) = MeasureFreq
                StoreEachIndex = StoreEachIndex + 1

                If j = SrcEndBit Then
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq_F1.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureFreq_F2.Math.Subtract(TrimTarget).Abs)
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                    PastDSPWave = InDSPwave
                Else
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                    PastDSPWave = InDSPwave
                End If

                TestLimitIndex = TestLimitIndex + 1

                '' 20160712 - Modify to use WriteComment to display output frequency.

                If gl_Disable_HIP_debug_log = False Then

                    For Each site In TheExec.sites
        ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        If b_KeepGoing(site) = True Then
                            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.pins(0).value(site) / 1000000), 6) & "M Hz")
                        End If
        ''                End If
                    Next site
                End If

                ''20160716 - Modify display info sequence when source bit in the section end
                If j = SrcEndBit Then
                    For Each site In TheExec.sites
    ''                    If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then

                        If b_KeepGoing(site) = True And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 1 & " Output Frequency = " & FormatNumber((MeasureFreq_F1.pins(0).value(site) / 1000000), 6) & "M Hz")
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 2 & " Output Frequency = " & FormatNumber((MeasureFreq_F2.pins(0).value(site) / 1000000), 6) & "M Hz")
                        End If
    ''                    End If
                    Next site
                    LastSectionF1F2_Index = LastSectionF1F2_Index + 2
                End If

                '' 20160712 - Compare Measure Frequency whether match target Freq
                b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                            StoredTargetTrimCode(site) = InDSPwave(site).COPY
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                Next site
                ''20161128-Stop trim code process if found out match code of all site
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
            Next j
        Next i
    End If

    ''============================================================================
    ''20161128 Findout mimiumn trim code
'    Dim CloseTargetFreq As New PinListData
'    Dim DiffValue As New SiteLong, PreviousDiffValue As New SiteLong, CloseIndex As New SiteLong
'
'    Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteLong
'    Dim BINstr As String
'    Dim CloseTargetTrimCode As New DSPWave
'    Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
'    Dim b_FirstTimeSwitch As Boolean
'    TheExec.sites.Selected = True
    If b_KeepGoing.All(False) Then
    Else
''        If TrimPrcocessAll = True Then
            CloseTargetTrimCode.CreateConstant 0, l_DigSrc_Sample_Size, DspLong

            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If StoredTargetTrimCode(site).CalcSum = 0 Then
                        b_UseMinTrim(site) = True
                    End If
                End If
            Next site

            If UseMinimumTrimCode = True Then
                b_UseMinTrim = True
            End If
            ''TEST
'            TheExec.sites.Selected = True
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_UseMinTrim(site) = True Then
                        '' Findout minimum difference value
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            DiffValue(site) = Abs(StoreEachTrimFreq(i).pins(0).value(site) - TrimTarget)
                            If DiffValue(site) <= PreviousDiffValue(site) Then
                                CloseIndex(site) = i
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                            If i = 0 Then
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                        Next i
                        '' Transfer to decimal value to findout minimum code
                        PreviousDecVal = 0
                        DecVal = 0
                        b_FirstTimeSwitch = False
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            binstr = vbNullString
                            If Abs(StoreEachTrimFreq(i).pins(0).value(site) - TrimTarget) = MinDiffVal(site) Then
                                For j = 0 To StoreEachTrimCode(i)(site).SampleSize - 1
                                    binstr = binstr & StoreEachTrimCode(i)(site).Element(j)
                                Next j
                                DecVal = Bin2Dec_rev_Double(binstr)

                                If DecVal < PreviousDecVal Then
                                    MinDecVal = DecVal
                                    CloseTargetTrimCode(site) = StoreEachTrimCode(i)(site).COPY
                                End If
                                PreviousDecVal = DecVal
                                If b_FirstTimeSwitch = False Then
                                    CloseTargetTrimCode(site) = StoreEachTrimCode(i)(site).COPY
                                    b_FirstTimeSwitch = True
                                End If
                            End If
                        Next i
                    End If
                End If
            Next site
''        End If
    End If

    For Each site In TheExec.sites
        If b_KeepGoing(site) = True Then
            If b_UseMinTrim(site) = True Then
                StoredTargetTrimCode(site) = CloseTargetTrimCode(site).COPY
            Else
                StoredTargetTrimCode(site) = StoredTargetTrimCode(site).COPY
            End If
        Else
            StoredTargetTrimCode(site) = StoredTargetTrimCode(site).COPY
        End If
    Next site
    ''============================================================================

    ''
TrimCodeTTR:
    TheExec.sites.Selected = True
'    Dim ZeroData() As Long
'    ReDim ZeroData(DigSrc_Sample_Size - 1)
'
'
'    For Each Site In TheExec.sites
'        If bTrimCodeFreqContinue(Site) = False Then
'            StoredTargetTrimCode(Site).Data = ZeroData
'        End If
'    Next Site

    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, StoredTargetTrimCode)
    End If



    Call HardIP_WriteFuncResult(, , glb_TestInstance)

    For Each site In TheExec.sites
        OutputTrimCode = vbNullString
        For k = 0 To StoredTargetTrimCode(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(StoredTargetTrimCode(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Final Output Trim Code = " & OutputTrimCode)
    Next site

    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel StoredTargetTrimCode, l_DigSrc_Sample_Size, ConvertedDataWf

    TestNameInput = Report_TName_From_Instance("C", DigSrc_pin.value, "TrimCode(Decimal)", 0, 0)

    TheExec.flow.TestLimit ConvertedDataWf.Element(0), 0, 2 ^ l_DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow  'for Turks
    'TheExec.Flow.TestLimit ConvertedDataWf.Element(0), 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow


    'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, StoredTargetTrimCode)
    '/* ------ Added by Kaino on 2019/06/12 ------- */
    For Each site In TheExec.sites
        InDSPwave_Repeat = StoredTargetTrimCode.Repeat(TrimCodeRepeat)
                If gl_Disable_HIP_debug_log = False Then
                OutputTrimCode = vbNullString
                For k = 0 To InDSPwave_Repeat(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave_Repeat(site).Element(k))
                    If (k + 1) Mod l_DigSrc_Sample_Size = 0 And k < InDSPwave_Repeat(site).SampleSize - 1 Then
                    OutputTrimCode = OutputTrimCode & ","
                    End If
            Next k
                TheExec.Datalog.WriteComment ("Site_" & site & " Final Test Trim Code = " & OutputTrimCode)
                End If

    Next site
''    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
'        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'    Else
'        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
'    End If
    '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */

    If BurstYesPatDict.Exists(LCase(patset.value)) = True Then
        tempVarArray = TheHdw.DSSC.pins(DigSrc_pin).Pattern(tempPatArr(tempSrcPatSeq)).Source.Labels.list ''20210609 temp
        sSrcSigName = tempVarArray(0)
'        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
        Call SetupDigSrcDspWave(tempPatArr(tempSrcPatSeq), DigSrc_pin, sSrcSigName, l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
        Call TheHdw.patterns(tempPatArr(tempSrcPatSeq)).start
    Else
        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", l_DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
    Call TheHdw.patterns(PattArray(0)).start
    End If

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0

    For Each Ts In TestSequenceArray
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If

        ''Update Interpose_PreMeas 20170801
        ''20160923 - Add Interpose_PreMeas entry point by each sequence
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            Else 'Do nothing '20230601
            End If
        End If

        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
        End If

        ''Update Interpose_PreMeas 20170801
        ''20161206-Restore force condiction after measurement
        ''Call SetForceCondition("RESTORE")
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            Else 'Do nothing '20230601
            End If
        End If

        TestSeqNum = TestSeqNum + 1

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts

    TheHdw.Digital.Patgen.HaltWait

    If TPModeAsCharz_GLB Then
        Disable_Inst_pinname_in_PTR
        TheExec.flow.TestLimit resultVal:=MeasureFreq, unit:=unitHz, ForceResults:=tlForceFlow
        Enable_Inst_pinname_in_PTR
    Else
        For p = 0 To MeasureFreq.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("F", MeasureFreq.pins(p), "Final", CInt(p))
            TheExec.flow.TestLimit resultVal:=MeasureFreq, unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next p
    End If

    Dim sl_FUSE_Val As New SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else

    End If

    DebugPrintFunc patset.value

    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If

    TheHdw.Alarms.Check
    Exit Function

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_AP", "TrimCodeFreq_New_ALG") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function TMPS(patset As Pattern, CPUA_Flag_In_Pat As Boolean, DigSrc_pin As PinList, DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, DigSrc_Equation As String, DigSrc_Assignment As String, _
                           Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
                            Optional CUS_Str_DigSrcData As String, Optional CUS_Str_DigCapData As String = vbNullString, Optional TMPS_Fuse_string As String, Optional Validating_ As Boolean) As Long

'' Step 1 : trim code is 8 bit, show out measured volt and trimed code, target volt is 0.9v
'' Step 2 : start from 0x8 and add algorithm to decide +/- direction
'' while decimal < 2 ^ DigSrc_Sample_Size
'' convert decimal to binary reverse
'' input the binary reverse data to digSrc_assignment

'From T-Col
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    Dim x As Long
    Dim InDSPwave As New DSPWave
    Dim OutDspWave As New DSPWave
    Dim CapOut As String
    Dim SrcOut As String
    Dim site As Variant
    Dim pat As String
    Dim i As Integer
    Dim ShowDec As String
    Dim ShowOut As String
    Dim TrimBits As String
    Dim b_TestDone As Boolean
    Dim code(7) As Integer
    Dim SourceNum As Integer
    Dim k As Integer
    Dim CtrlBits As New DSPWave
    Dim MinValue As New DSPWave
    Dim Data_Array(7) As New SiteLong
    Dim PassFlag_TMPS As New SiteBoolean
    On Error GoTo errHandler

    b_TestDone = False
    SourceNum = 0
    CtrlBits.CreateConstant 0, 2 * DigSrc_Sample_Size / DigSrc_DataWidth
    MinValue.CreateConstant 0, DigSrc_Sample_Size / DigSrc_DataWidth

    If DigSrc_Sample_Size = 0 Then
        TheExec.Datalog.WriteComment ("Error!! - Please check input argument DigSrc_Sample_Size")
        Exit Function
    End If

    TheHdw.Digital.Patgen.Halt
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.patterns(patset).Load

    Dim PattArray() As String
    Dim PatCount As Long

    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode

    Do While b_TestDone = False
        For Each site In TheExec.sites.Active

          ''  theexec.Datalog.WriteComment ("======== Start Dig Src setup =======")
            If SourceNum = 0 Then
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, site)
            End If

            '' Processing Captured Code and Printing the Values
            If SourceNum > 0 Then
                For k = 0 To DigCap_Sample_Size / DigCap_DataWidth - 1
                code(k) = 0
                    For i = 0 To DigCap_DataWidth - 1
                        code(k) = code(k) + OutDspWave(site).Element(k * DigCap_DataWidth + i) * (2 ^ i)
                    Next i
                ''   Code(k) = 3770 + Int(Rnd * 20)    '' Random Codes for Offline testing
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Capture data_ " & k & "_Value =" & code(k)
                Next k
            End If

            If SourceNum > 0 And SourceNum < 6 Then
                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                    If code(k) > 3780 Then
                        For x = 0 To 3
                            InDSPwave(site).Element(k * 28 + 4 * (7 - SourceNum) + x) = 0
                        Next x
                    Else
                        For x = 0 To 3
                            InDSPwave(site).Element(k * 28 + 4 * (7 - SourceNum) + x) = 1
                        Next x
                    End If

                    If SourceNum < 5 Then
                        For x = 0 To 3
                            InDSPwave(site).Element(k * 28 + 4 * (6 - SourceNum) + x) = 1
                        Next x
                    End If
                Next k

            End If

            ''' Trim Control(first 2) Bits

            If SourceNum > 4 And SourceNum < 9 Then
                Select Case SourceNum - 5

                    Case 0
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            For x = 0 To 7
                                InDSPwave(site).Element(k * 28 + x) = 0
                            Next x
                        Next k


                    Case 1
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            MinValue(site).Element(k) = code(k)
                            For i = 0 To 15
                                CtrlBits(site).Element(i) = 0
                            Next i

                            For x = 0 To 3
                                InDSPwave(site).Element(k * 28 + x) = 0
                            Next x
                            For x = 4 To 7
                                InDSPwave(site).Element(k * 28 + x) = 1
                            Next x
                        Next k


                    Case 2
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            If code(k) < MinValue(site).Element(k) Then
                                MinValue(site).Element(k) = code(k)
                                CtrlBits(site).Element(2 * k) = 0
                                CtrlBits(site).Element(2 * k + 1) = 1
                            End If

                            For x = 0 To 3

                                InDSPwave(site).Element(k * 28 + x) = 1
                            Next
                            For x = 4 To 7
                                InDSPwave(site).Element(k * 28 + x) = 0
                            Next x
                        Next k


                    Case 3
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            If code(k) < MinValue(site).Element(k) Then
                                MinValue(site).Element(k) = code(k)
                                CtrlBits(site).Element(2 * k) = 1
                                CtrlBits(site).Element(2 * k + 1) = 0
                            End If

                            For x = 0 To 7
                                InDSPwave(site).Element(k * 28 + x) = 1
                            Next x
                        Next k

                End Select
            End If


            ''''' Sourcing Trimmed Data Bits

            If SourceNum = 9 Then

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                    If code(k) < MinValue(site).Element(k) Then
                        MinValue(site).Element(k) = code(k)
                        CtrlBits(site).Element(2 * k) = 1
                        CtrlBits(site).Element(2 * k + 1) = 1
                    End If

                Next k

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1
                    For x = 0 To 3
                        InDSPwave(site).Element(k * 28 + x) = CtrlBits(site).Element(2 * k)
                    Next x

                    'InDSPWave(site).Element(k * 28 + 1) = CtrlBits(site).Element(2 * k)
                    'InDSPWave(site).Element(k * 28 + 2) = CtrlBits(site).Element(2 * k)
                    'InDSPWave(site).Element(k * 28 + 3) = CtrlBits(site).Element(2 * k)
                    For x = 4 To 7
                        InDSPwave(site).Element(k * 28 + x) = CtrlBits(site).Element(2 * k + 1)
                    Next x
                    'InDSPWave(site).Element(k * 28 + 5) = CtrlBits(site).Element(2 * k + 1)
                    'InDSPWave(site).Element(k * 28 + 6) = CtrlBits(site).Element(2 * k + 1)
                    'InDSPWave(site).Element(k * 28 + 7) = CtrlBits(site).Element(2 * k + 1)
                Next k

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1
                    Data_Array(k) = 0
                    For i = 0 To DigSrc_DataWidth / 4 - 1
                        Data_Array(k) = Data_Array(k) + InDSPwave(site).Element(k * DigSrc_DataWidth + 4 * i) * (2 ^ i)
                    Next i
                Next k

            End If



        Next site


        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "Meas_Src", DigSrc_Sample_Size, InDSPwave)

        If SourceNum = 9 Then
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "==========SOURCING TRIMMED DATA BITS==============="
        End If

        For Each site In TheExec.sites.Active
            If SourceNum > 0 Then
                SrcOut = vbNullString
                For i = 0 To DigSrc_Sample_Size - 1
                    SrcOut = SrcOut & InDSPwave(site).Element(i)
                    If i Mod DigSrc_DataWidth = DigSrc_DataWidth - 1 Then
                        SrcOut = SrcOut & ", "
                    ElseIf i Mod 4 = 3 Then
                        SrcOut = SrcOut & " "
                    End If

                Next i
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Source Data=" & SrcOut
            End If
        Next site

        If DigCap_Sample_Size <> 0 Then

           TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test Start ========")
           Call DigCapSetup(PattArray(0), DigCap_Pin, "Meas_cap", DigCap_Sample_Size, OutDspWave)

        End If

        SourceNum = SourceNum + 1

        Call TheHdw.patterns(PattArray(0)).start

        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end

        PatCount = PatCount + 1
        CapOut = vbNullString


        TheHdw.Digital.Patgen.HaltWait   '' Haltwait at patten end

        If SourceNum = 10 Then
            b_TestDone = True
        End If

    Loop

    For Each site In TheExec.sites.Active
        For k = 0 To DigCap_Sample_Size / DigCap_DataWidth - 1
        code(k) = 0
            For i = 0 To DigCap_DataWidth - 1
                code(k) = code(k) + OutDspWave(site).Element(k * 12 + i) * (2 ^ i)
            Next i
         If CurrentJobName_L Like "*char*" Then
         Disable_Inst_pinname_in_PTR
            TheExec.flow.TestLimit resultVal:=code(k), unit:=unitNone, ForceResults:=tlForceFlow
         Enable_Inst_pinname_in_PTR
         Else
            TheExec.flow.TestLimit resultVal:=code(k), unit:=unitNone, Tname:="Capture Data_" & k, ForceResults:=tlForceFlow
            End If
        Next k

'' Add for TMPS fusing 20160526
       If TheHdw.Digital.Patgen.PatternBurstPassed(site) = False Then 'Pattern Fail
           PassFlag_TMPS(site) = False
       Else
           PassFlag_TMPS(site) = True
        End If
    Next site
    
    ''20210406 Modify for newEfuse
    Dim opbank As eFuseBdfBank
    Dim opbank_cfg As eFuseBdfBank, opbank_mon As eFuseBdfBank, opbank_udr As eFuseBdfBank
    Dim Pass_Fail_Flag As New SiteBoolean: Pass_Fail_Flag = True
    Dim Data_Array_Default As New SiteLong: Data_Array_Default = 64
    
    If LCase(TMPS_Fuse_string) Like "*t3_fuse_25c*" Then
        For x = 1 To 7
            Set opbank = GetBdfBank("ECID")
            opbank.SetEfuse "Temp" & x, Data_Array(x), PassFlag_TMPS, , , , True
        Next x
    ElseIf LCase(TMPS_Fuse_string) Like "*t5_fuse_85c*" Then
        Set opbank_cfg = GetBdfBank("CFG")
        Set opbank_mon = GetBdfBank("MON")
        Set opbank_udr = GetBdfBank("UDR")
        
        If PassFlag_TMPS(site) = True Then
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL0", Data_Array(0), Pass_Fail_Flag, , , , True
            opbank_mon.SetEfuse "THERMAL_PARAM_TS_REFERENCE_CTRL", Data_Array(0), Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL1", Data_Array(1), Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL2", Data_Array(2), Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL3", Data_Array(3), Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor3_tTRIM", Data_Array(4), Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor2_tTRIM", Data_Array(5), Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor1_tTRIM", Data_Array(6), Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor0_tTRIM", Data_Array(7), Pass_Fail_Flag, , , , True
        Else
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL0", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_mon.SetEfuse "THERMAL_PARAM_TS_REFERENCE_CTRL", Data_Array(0), Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL1", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL2", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL3", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor3_tTRIM", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor2_tTRIM", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor1_tTRIM", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor0_tTRIM", Data_Array_Default, Pass_Fail_Flag, , , , True
        End If
    End If

    Call HardIP_WriteFuncResult

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in TMPS function"
    If AbortTest Then Exit Function Else Resume Next

End Function


