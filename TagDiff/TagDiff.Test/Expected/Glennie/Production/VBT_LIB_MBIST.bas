Attribute VB_Name = "VBT_LIB_MBIST"
Option Explicit

Public gL_1st_MbistSheetRead As Long ''''to read Mbist Flow ChkList sheet

Public Const gS_CpuMbist_sheetName = "CpuMbist_Flow_ChkList"
Public Const gS_GpuMbist_sheetName = "GpuMbist_Flow_ChkList"
Public Const gS_SocMbist_sheetName = "SocMbist_Flow_ChkList"
Public gS_Mbist_GroupName As String
Public gL_Mbist_GroupName_Index As Long
Public gS_currPayload_pattSetName As String ''''used to record the latest patload pattSetName
Public gB_enable_NewMbist_flag As Boolean

Public DebugPrtImm As Boolean
Public DebugPrtDlog As Boolean
'////////////////////////////////////////////////////////////////////////for mbist loop module
Private Type patlist
   pat() As String
   instance() As String
End Type

Private Type timinglevels
    dc_spec As String
    dc_sel As String
    ac_spec As String
    ac_sel As String
    timingsheet As String
    levelsheet As String
End Type

Private Type mbist_inf_list
    block_name As String
    block_count As Long
    block_count_name() As String
    ins_name As String
    pat_name() As String
    pat_tested() As Boolean
    pat_count As Long
    
    Flag_Name() As String
    flag_count As Long
    
    timing_levels As timinglevels
    GroupIdentify() As String
    block_type_pat() As patlist
End Type

Private Type mbist_inf
    Block() As mbist_inf_list
End Type

Public Mbist(2) As mbist_inf

'''''''''''''''''''''''''''''''''''
Private Type mbist_match_inf_list
    binflag_match_name As String
    binflag_mid_name As String
End Type

Private Type mbist_match_inf
    inst_nu() As mbist_match_inf_list
    inst_count As Long
End Type


Public mbist_match(2) As mbist_match_inf
'''''''''''''''''''''''''''''''''''
Public Const CPU_sheet = 1
Public Const SOC_sheet = 0
Public File_path As String
'///////////////////////////////////////////////////////////
'=========================================================
Private Type patlist_dynamic
   pat_dynamic() As String
   instance_dynamic() As String
End Type

Private Type mbist_inf_list_dynamic
    block_name_dynamic As String
    block_count_dynamic As Long
    block_count_name_dynamic() As String
    ins_name_dynamic As String
    pat_name_dynamic() As String
    pat_tested_dynamic() As Boolean
    pat_count_dynamic As Long
    
    flag_name_dynamic() As String
    flag_count_dynamic As Long
    
    block_type_pat_dynamic() As patlist_dynamic
End Type

Private Type mbist_inf_dynamic
    Block_dynamic() As mbist_inf_list_dynamic
End Type

Public mbist_dynamic As mbist_inf_dynamic

Public type_nu As Double
Public confirm_in_loop As Boolean
Public create_flag_sheet As Boolean
Public create_flag_begin As Boolean

Public index_flag_x As Long
Public index_flag_y As Long

Public bist_type As String

Public mbist_flag_set_placement As Long   ''20180928 block loop flag


Public Function Protect_Mbist_Sheet()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "Protect_Mbist_Sheet"

    Sheets(gS_CpuMbist_sheetName).Protect
    ''Sheets(gS_GpuMbist_sheetName).Protect
    ''Sheets(gS_SocMbist_sheetName).Protect
    
    '========================================
    '=  How to unprotect the spread sheet   =
    '========================================
    'Step1.
        'Select the spread sheet of gS_CpuMbist_sheetName
    'Step2.
        'Switch the VBT window and launch the Immediate Window
    'Step3.
        'In the Immediate Window, key in the following command
        '===>  activesheet.unprotect
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function UnProtect_Mbist_Sheet()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "UnProtect_Mbist_Sheet"

    Sheets(gS_CpuMbist_sheetName).Unprotect
    ''Sheets(gS_GpuMbist_sheetName).Unprotect
    ''Sheets(gS_SocMbist_sheetName).Unprotect

    '========================================
    '=  How to unprotect the spread sheet   =
    '========================================
    'Step1.
        'Select the spread sheet of gS_CpuMbist_sheetName
    'Step2.
        'Switch the VBT window and launch the Immediate Window
    'Step3.
        'In the Immediate Window, key in the following command
        '===>  activesheet.unprotect
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_Mbist_SetLoopCNT_BM(bistType As String, block_type As String, Optional block_loop_str As String = vbNullString, Optional flag_long As Long = 7) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Mbist_SetLoopCNT_BM"

    Dim site As Variant
    Dim m_tsname As String
    Dim g As Long, p As Long, k As Long, S As Long:: g = 0:: p = 0:: k = 0:: S = 0
    Dim nu As Long, nu01 As Long, nu02 As Long:: nu = 0:: nu01 = 0:: nu02 = 0
    Dim block_loop_name() As String
    Dim block_loop_sum As Long

    Dim conbine_string() As String
    ReDim conbine_string(UBound(Mbist(type_nu).Block)) As String
    Dim pre_type As String
    
    mbist_flag_set_placement = flag_long
    
    confirm_in_loop = True   ''check whether in the mbist loop
    bist_type = bistType
    '===================================================
    If (bistType = "CPU") Then
        type_nu = 1
    ElseIf (bistType = "SOC") Then
        type_nu = 0
    End If

    For g = 0 To UBound(Mbist(type_nu).Block)
        For Each site In TheExec.sites
            TheExec.sites(site).SiteVariableValue("LP_BM") = 0
        Next site
    Next g
    '===================================================
    If (block_loop_str = "") Then
        For Each site In TheExec.sites
            For g = 0 To UBound(Mbist(type_nu).Block)
                If (Mbist(type_nu).Block(g).block_name = block_type) Then
                    TheExec.sites(site).SiteVariableValue("LCount_BM") = UBound(Mbist(type_nu).Block(g).block_type_pat) + 1
                End If
            Next g
        Next site
    Else
        block_loop_name = Split(block_loop_str, ",")
        
        For Each site In TheExec.sites
             TheExec.sites(site).SiteVariableValue("LCount_BM") = UBound(block_loop_name) + 1
        Next site
    End If
    '===================================================dynamic init
     ReDim mbist_dynamic.Block_dynamic(0)
    '===================================================transfer dynamic assign
    ReDim Preserve mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(50)
    ReDim Preserve mbist_dynamic.Block_dynamic(0).pat_name_dynamic(50)
    ReDim Preserve mbist_dynamic.Block_dynamic(0).pat_tested_dynamic(50)
    mbist_dynamic.Block_dynamic(0).block_count_dynamic = 0
    mbist_dynamic.Block_dynamic(0).ins_name_dynamic = vbNullString
    mbist_dynamic.Block_dynamic(0).pat_count_dynamic = 0
    For g = 0 To UBound(Mbist(type_nu).Block)
        For k = 0 To Mbist(type_nu).Block(g).block_count - 1
            If (Mbist(type_nu).Block(g).block_name = block_type) Then
                ReDim Preserve mbist_dynamic.Block_dynamic(0).pat_name_dynamic(S)
                ReDim Preserve mbist_dynamic.Block_dynamic(0).pat_tested_dynamic(S)
                If (block_loop_str = "") Then
                        'mbist_dynamic.Block_dynamic(0).block_count_dynamic = Mbist(type_nu).Block(g).block_count
                        If (pre_type <> Mbist(type_nu).Block(g).block_count_name(k) And pre_type <> "") Then
                            p = p + 1
                        End If
                        mbist_dynamic.Block_dynamic(0).block_count_dynamic = p + 1
                        ReDim Preserve mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(p)
                        mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(p) = Mbist(type_nu).Block(g).block_count_name(k)
                        pre_type = Mbist(type_nu).Block(g).block_count_name(k)

                        mbist_dynamic.Block_dynamic(0).block_name_dynamic = Mbist(type_nu).Block(g).block_name
                        mbist_dynamic.Block_dynamic(0).ins_name_dynamic = Mbist(type_nu).Block(g).ins_name
                        mbist_dynamic.Block_dynamic(0).pat_count_dynamic = Mbist(type_nu).Block(g).pat_count
                        mbist_dynamic.Block_dynamic(0).pat_name_dynamic(S) = Mbist(type_nu).Block(g).pat_name(k)
                        mbist_dynamic.Block_dynamic(0).pat_tested_dynamic(S) = False
                        S = S + 1
                        '--------------------------------------------------------------------------
                        For nu = 0 To UBound(Mbist(type_nu).Block(g).block_type_pat)
                            For nu01 = 0 To UBound(Mbist(type_nu).Block(g).block_type_pat(nu).instance)
                                ReDim Preserve mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu)
                                ReDim Preserve mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu)
                                ReDim Preserve mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu).instance_dynamic(nu01)
                                ReDim Preserve mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu).pat_dynamic(nu01)
                                mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu).instance_dynamic(nu01) = Mbist(type_nu).Block(g).block_type_pat(nu).instance(nu01)
                                mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu).pat_dynamic(nu01) = Mbist(type_nu).Block(g).block_type_pat(nu).pat(nu01)
                            Next nu01
                        Next nu
                        '--------------------------------------------------------------------------
                Else
                    If (UBound(block_loop_name) >= p) Then
                        If (block_loop_name(p) Like Mbist(type_nu).Block(g).block_count_name(k)) Then
                            ReDim Preserve mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(p)
                            mbist_dynamic.Block_dynamic(0).block_count_dynamic = mbist_dynamic.Block_dynamic(0).block_count_dynamic + 1
                            mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(p) = Mbist(type_nu).Block(g).block_count_name(k)
                            mbist_dynamic.Block_dynamic(0).block_name_dynamic = Mbist(type_nu).Block(g).block_name
                            mbist_dynamic.Block_dynamic(0).ins_name_dynamic = Mbist(type_nu).Block(g).ins_name
                            'p = p + 1
                            '--------------------------------------------------------------------------
                            For nu = 0 To UBound(Mbist(type_nu).Block(g).block_type_pat)
                                For nu01 = 0 To UBound(Mbist(type_nu).Block(g).block_type_pat(nu).instance)
                                    If (UCase(Mbist(type_nu).Block(g).block_type_pat(nu).pat(nu01)) Like UCase("*_" + block_loop_name(p) + "_*") Or UCase(Mbist(type_nu).Block(g).block_type_pat(nu).pat(nu01)) Like UCase("*_" + block_loop_name(p))) Then
                                        If (pre_type <> block_loop_name(p) And pre_type <> "") Then
                                            nu02 = nu02 + 1
                                        End If
                                        ReDim Preserve mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu02)
                                        ReDim Preserve mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu02)
                                        ReDim Preserve mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu02).instance_dynamic(nu01)
                                        ReDim Preserve mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu02).pat_dynamic(nu01)
                                        mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu02).instance_dynamic(nu01) = Mbist(type_nu).Block(g).block_type_pat(nu).instance(nu01)
                                        mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(nu02).pat_dynamic(nu01) = Mbist(type_nu).Block(g).block_type_pat(nu).pat(nu01)
                                        pre_type = block_loop_name(p)
                                    End If
                                Next nu01
                            Next nu
                            '--------------------------------------------------------------------------
                            p = p + 1
                        End If
                    End If
                    If (UBound(block_loop_name) >= S) Then
                        If (Mbist(type_nu).Block(g).pat_name(k) Like UCase("*_" + block_loop_name(S)) + "*") Then
                            mbist_dynamic.Block_dynamic(0).pat_count_dynamic = mbist_dynamic.Block_dynamic(0).pat_count_dynamic + 1
                            mbist_dynamic.Block_dynamic(0).pat_name_dynamic(S) = Mbist(type_nu).Block(g).pat_name(k)
                            mbist_dynamic.Block_dynamic(0).pat_tested_dynamic(S) = False
                            S = S + 1
                        End If
                    End If
                End If
            End If
        Next k
    Next g
    p = 0:: S = 0

    index_flag_y = 1
    index_flag_x = index_flag_x + 1

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Mbist_Initialize()
    On Error GoTo errHandler
        Dim FileExists As Boolean
        
        '---------------------------------------------------- for mbist loop module
        If mbist_sheet_init <> True Then
            If TheExec.enableWord("Mbist_FingerPrint") = True Then
                Call Init_MBISTFailBlock
            End If
            init_MBIST_ChkList_block_loop ("SOC")
            'init_MBIST_ChkList_block_loop ("CPU")
            
            ''' preload pattern
            Dim pat As Variant
            Dim pats As Variant
            Dim i As Integer
            Dim j As Integer
            Dim k As Integer
            Dim path_pat() As String
            Dim path_cnt As Long

            For i = 0 To UBound(Mbist) - 2          '------------------------------"-2 is Hard Coding for SOC Only
                For j = 0 To UBound(Mbist(i).Block)
                    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered, , , , _
                                                        Mbist(i).Block(j).timing_levels.levelsheet, _
                                                        Mbist(i).Block(j).timing_levels.dc_spec, _
                                                        Mbist(i).Block(j).timing_levels.dc_sel, _
                                                        Mbist(i).Block(j).timing_levels.timingsheet, _
                                                        Mbist(i).Block(j).timing_levels.ac_spec, _
                                                        Mbist(i).Block(j).timing_levels.ac_sel)
                    For k = 0 To UBound(Mbist(i).Block(j).pat_name)
                        path_pat = TheExec.DataManager.Raw.GetPatternsInSet(Mbist(i).Block(j).pat_name(k), path_cnt)
                        Mbist(i).Block(j).pat_name(k) = path_pat(0)
                    Next k
                    Call TheHdw.patterns(Join(Mbist(i).Block(j).pat_name, ",")).Load
'                    For Each pat In Mbist(i).block(j).pat_name
'                        Call thehdw.Patterns(pat).ValidatePatlist
'                        Dim patCnt As Long
'                        Dim patSet() As String
'                        patSet = TheExec.DataManager.Raw.GetPatternsInSet(pat, patCnt)
'
'                        For Each pats In patSet
'                            Call thehdw.Patterns(pats).ValidatePatlist
'                        Next pats
'                    Next pat
                Next j
            Next i
            
            mbist_sheet_init = True
        End If
        
        create_flag_sheet = False  '//print flag list
        index_flag_y = 0
        index_flag_x = 0
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''delet file
        File_path = ".\Mbist_Block_loop_flag_list.csv"
        FileExists = (Dir(File_path) <> "")
        If FileExists Then
               SetAttr File_path, vbNormal
               Kill File_path
        End If
        '----------------------------------------------------
    
    
    Exit Function
errHandler:
   If AbortTest Then Exit Function Else Resume Next
End Function

Public Function init_MBIST_ChkList_block_loop(bistType As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "init_MBIST_ChkList_block_loop"

    Dim sheetName As String
    Dim MySheet As Worksheet
    Dim myCell As Object
    Dim groupCell As Object
    Dim pre_groupCell As Object
    Dim offCell As Object
    Dim myCell_Header As Object

    Dim myCellA1 As Object
    Dim m_A1_rowCnt As Long

    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim n As Long
    Dim p As Long
    Dim g As Long
    Dim pre_store_nu As Long
    Dim store_nu As Long
    Dim store_string As String
    Dim findout_inst As Boolean
    
    Dim m_cellCnt As Long
    Dim m_cellStr As String
    Dim m_groupCell As String
    Dim m_pre_groupCell As String
    Dim BlockCounter As Long
    Dim m_cellStr2 As String
    Dim m_offcolStr As String
    Dim m_lastrow As Long
    Dim m_lastNCnt As Long

    Dim find_1stHeader As Boolean
    Dim FIND_AllHEADER As Boolean

    Dim idx_END As Long
    Dim idx_Instance_BM_Decision_row As Long
    Dim idx_Instance_BM_Decision_column As Long
    Dim idx_BM_Pattern_column As Long
    Dim idx_Block_Name_column As Long
    Dim idx_Pattern_Block_column As Long

    Dim idx_BinFlag_Decision_column As Long
    Dim idx_Name_for_BinFlag_column As Long
    Dim idx_BinFlag_Mid_Name_column As Long
    Dim idx_BinFlag_with_PM_BM_column As Long
    
    Dim idx_DCSpec_column As Long
    Dim idx_DCSel_column As Long
    Dim idx_ACSpec_column As Long
    Dim idx_ACSel_column As Long
    Dim idx_TimingSheet_column As Long
    Dim idx_LevelSheet_column As Long
    Dim idx_GroupIdentify As Long

    Dim m_cnt As Long
    Dim m_idx As Long
    Dim m_pattname As String
    Dim m_pattRawname As String
    Dim m_debug_PMode As String
    Dim m_debug_Block As String

    Dim m_InsCnt As Long
    Dim m_BM_pat_Cnt As Long

    Dim m_patArr() As String
    Dim m_patcount As Long
    Dim m_Block_Cnt As Long
    Dim m_Block_cata_Cnt As Long
    Dim m_Pat_Cnt As Long

    Dim number As Long
    Dim Character As String
    Dim pre_Character As String
    Dim counter As Long
    Dim counter01 As Long
    Dim sheet_type As Long
    Dim confirm_type As Long
    
    Dim block_type As String
    Dim pre_block_type As String
    
    
    Dim Loc_dash As Integer
    ''''''''''''''''''''''''''''''''''''
    Dim instance_name() As String
    ReDim instance_name(100)
    Dim ins_true_name() As String
    ReDim ins_true_name(100)

    Dim instance_flag() As Boolean
    ReDim instance_flag(100)
    Dim pattern_name() As String
    ReDim pattern_name(100)
    Dim block_name() As String
    ReDim block_name(100)
    Dim block_name_count() As Long
    ReDim block_name_count(100)

    Dim block_count_name() As String
    ReDim block_count_name(100, 100)

'''    Dim block_type_pat() As String
'''    ReDim block_type_pat(50, 50, 50)
    
    Dim inst_group_name() As String
    ReDim inst_group_name(100)
    Dim flag_pm_bm() As Boolean
    ReDim flag_pm_bm(100)

    Dim flag_match_name() As String
    ReDim flag_match_name(100)
    Dim flag_mid_name() As String
    ReDim flag_mid_name(100)

    Dim block_nu As Double
    Dim block_type_nu As Double
    Dim pat_number As Double

    Dim m_flag_range As Long
    m_flag_range = 0
    ''''-------------------------------
    DebugPrtImm = False
    DebugPrtDlog = False
    ''''-------------------------------
    counter = 0
    bistType = UCase(bistType)
    gB_findPwrPin_flag = False ''''Initial
    '===================================================
    If (bistType = "CPU") Then
        gB_findCpuMbist_flag = False
        sheetName = gS_CpuMbist_sheetName
        sheet_type = CPU_sheet
    ElseIf (bistType = "GPU") Then

    ElseIf (bistType = "SOC") Then
        gB_findSocMbist_flag = False
        sheetName = gS_SocMbist_sheetName
        sheet_type = SOC_sheet
    End If
    '===================================================
    Set MySheet = Sheets(sheetName)
    Set myCellA1 = MySheet.range("A1")
    m_A1_rowCnt = 0
    Set myCell = MySheet.range("A1")
    m_cellStr = UCase(Trim(myCell.value))

    find_1stHeader = False
    FIND_AllHEADER = False
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Do
        '======================================1......To find the 1st Word "Test*Instance*Decision"
        'Here Search Cell from Left to Right
        If (find_1stHeader = False) Then
            m_cellCnt = 0
            Do While (m_cellCnt < 5)
                If (m_cellStr Like UCase("Test*Instance*Decision")) Then
                    idx_Instance_BM_Decision_row = myCell.Row          '//y
                    idx_Instance_BM_Decision_column = myCell.Column    '//x
                    find_1stHeader = True
                    Exit Do
                End If
                Set myCell = myCell.offset(rowOffset:=0, columnOffset:=1)
                m_cellStr = UCase(Trim(myCell.value))
                m_cellCnt = m_cellCnt + 1
            Loop
        End If

        '======================================2......To find the following Header Words
        'By each Header, get the related parameters.
        If (find_1stHeader) Then
            If (bistType = "CPU") Then
                Set myCell = myCellA1.offset(rowOffset:=idx_Instance_BM_Decision_row - 1, columnOffset:=0)
            ElseIf (bistType = "GPU") Then
            ElseIf (bistType = "SOC") Then
                Set myCell = myCellA1.offset(rowOffset:=idx_Instance_BM_Decision_row - 1, columnOffset:=0)
            End If
            m_cellStr = UCase(Trim(myCell.value))  '//remove sapce(front and end)
            m = 0
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            Do While (UCase(m_offcolStr) <> "END")
                Set offCell = myCell.offset(rowOffset:=0, columnOffset:=m) ''''search cell from Left to Right
                m_offcolStr = UCase(Trim(offCell.value))
                m = m + 1
                ''''-------------------------
                ''''Column Sequence
                ''''-------------------------
                If (m_offcolStr Like UCase("Test*Instance*Decision")) Then
                    idx_Instance_BM_Decision_column = m
                ElseIf (m_offcolStr Like UCase("*BM*Pattern*")) Then
                    idx_BM_Pattern_column = m
                ElseIf (m_offcolStr Like UCase("Block*Name*")) Then
                    idx_Block_Name_column = m
                ElseIf (m_offcolStr Like UCase("*Pattern_Block*")) Then
                    idx_Pattern_Block_column = m
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf (m_offcolStr Like UCase("*BinFlag*Decision*Type*")) Then
                    idx_BinFlag_Decision_column = m
                ElseIf (m_offcolStr Like UCase("*Match*Name*for*BinFlag*")) Then
                    idx_Name_for_BinFlag_column = m
                ElseIf (m_offcolStr Like UCase("*BinFlag*Mid*Name*")) Then
                    idx_BinFlag_Mid_Name_column = m
                ElseIf (m_offcolStr Like UCase("*BinFlag*with*PM/BM*")) Then
                    idx_BinFlag_with_PM_BM_column = m
                ElseIf (m_offcolStr Like UCase("DC_Spec")) Then
                    idx_DCSpec_column = m
                ElseIf (m_offcolStr Like UCase("DC_Sel")) Then
                    idx_DCSel_column = m
                ElseIf (m_offcolStr Like UCase("AC_Spec")) Then
                    idx_ACSpec_column = m
                ElseIf (m_offcolStr Like UCase("AC_Sel")) Then
                    idx_ACSel_column = m
                ElseIf (m_offcolStr Like UCase("TimingSheet")) Then
                    idx_TimingSheet_column = m
                ElseIf (m_offcolStr Like UCase("LevelSheet")) Then
                    idx_LevelSheet_column = m
                ElseIf (m_offcolStr Like UCase("GroupIdentify")) Then
                    idx_GroupIdentify = m
                End If
            Loop ''''end of Do While (m_offcolStr <> "END")
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            If (m_offcolStr = ("END")) Then
                idx_END = m
                FIND_AllHEADER = True
            End If
        End If

        m_A1_rowCnt = m_A1_rowCnt + 1
        Set myCell = myCellA1.offset(rowOffset:=m_A1_rowCnt, columnOffset:=0)
        m_cellStr = UCase(Trim(myCell.value))
    Loop While (FIND_AllHEADER = False)

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If (FIND_AllHEADER) Then
        If (bistType = "CPU") Then
            Set myCell_Header = myCellA1.offset(rowOffset:=idx_Instance_BM_Decision_column - 1, columnOffset:=0)
        ElseIf (bistType = "GPU") Then
        ElseIf (bistType = "SOC") Then
            Set myCell_Header = myCellA1.offset(rowOffset:=idx_Instance_BM_Decision_column - 1, columnOffset:=0)
        End If
        m_cellStr = UCase(Trim(myCell_Header.value))
        'DebugPrintLog "4...(find_AllHeader=True) Row=" & myCell_Header.Row & ", Column=" & myCell_Header.Column & ", Cell=" & myCell_Header.Value & " (m_cellStr=" + m_cellStr + ")"

        ''''initialize -----------------------------------------
        If (bistType = "CPU") Then
             ReDim Mbist(sheet_type).Block(100)
        ElseIf (bistType = "GPU") Then
        ElseIf (bistType = "SOC") Then
             ReDim Mbist(sheet_type).Block(100)
        End If
        ''''----------------------------------------------------
        ''''Then get the following parameter per Header
        m = 0
        Do While (m <= idx_END)
            m = m + 1   ''''Column direction
            n = 0       ''''index and row direction

            Set myCell = myCell_Header.offset(rowOffset:=0, columnOffset:=(m - 1)) ''''rowOffset MUST be always '0'
            m_cellStr = UCase(Trim(myCell.value))
            'm_lastrow = myCell.End(xlDown).Row    '//y range
            m_lastrow = MySheet.Cells(Rows.Count, m).End(xlUp).Row
            If (bistType = "CPU") Then
                m_lastNCnt = m_lastrow - idx_Instance_BM_Decision_row
            ElseIf (bistType = "GPU") Then
            ElseIf (bistType = "SOC") Then
                m_lastNCnt = m_lastrow - idx_Instance_BM_Decision_row
            End If

            Select Case (m)
            Case idx_END
                Exit Do     ''''end
            '=================================================================================Test Instance Name for PM/BM Decision
            Case idx_Instance_BM_Decision_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    If (bistType = "CPU") Then
                        instance_name(n) = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                        instance_name(n) = m_cellStr
                    End If
                    n = n + 1
                Loop

                m_InsCnt = n
                ReDim Preserve instance_name(n - 1)    'redim and hold orignal data
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================Choose PM/BM Pattern
             Case idx_BM_Pattern_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    If (m_cellStr = "BM") Then
                       instance_flag(n) = True
                       ins_true_name(counter) = instance_name(n)
                       counter = counter + 1
                    Else
                       instance_flag(n) = False
                    End If
                    n = n + 1
                Loop

                ReDim Preserve instance_flag(n - 1)
                ReDim Preserve ins_true_name(counter - 1)
                m_BM_pat_Cnt = n
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================Block Name
             Case idx_Block_Name_column
                BlockCounter = 0
                Do While (n < m_lastNCnt)
                    number = 0: Character = vbNullString: block_type = vbNullString
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    
                    Dim tempcell As Object
                    Set tempcell = myCell
                    Set groupCell = tempcell.offset(rowOffset:=0, columnOffset:=8)
                    m_groupCell = UCase(Trim(groupCell.value))
                    Set pre_groupCell = groupCell.offset(rowOffset:=-1, columnOffset:=0) '///
                    m_pre_groupCell = UCase(Trim(pre_groupCell.value))
                    
                    block_type = m_cellStr

                    Call Separate_nu_char(m_cellStr, number, Character)
                    
                    
                    If pre_block_type = block_type Or pre_block_type = "" Then
                    
                    ReDim Preserve Mbist(sheet_type).Block(i).block_type_pat(BlockCounter)
                    Else
                    BlockCounter = BlockCounter + 1
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    n = n + 1
                    
                    If (pre_Character <> Character) Or (m_groupCell <> m_pre_groupCell) Then
                        block_name(counter01) = Character
                        block_count_name(counter01, block_name_count(counter01)) = m_cellStr
                        block_name_count(counter01) = block_name_count(counter01) + 1
                        BlockCounter = 0
                        If (pre_Character <> "") Then
                            counter = counter + 1
                        End If
                        counter01 = counter01 + 1
                        pre_Character = Character
                    Else
                        block_count_name(counter01 - 1, block_name_count(counter01 - 1)) = m_cellStr
                        block_name_count(counter01 - 1) = block_name_count(counter01 - 1) + 1
                    End If
                    
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If (pre_block_type = m_cellStr Or pre_block_type = "") Then
                        ReDim Preserve Mbist(sheet_type).Block(counter).block_type_pat(BlockCounter).pat(p)
                        Mbist(sheet_type).Block(counter).block_type_pat(BlockCounter).pat(p) = myCell.offset(rowOffset:=0, columnOffset:=1)
                        pre_block_type = m_cellStr
                        
                        p = p + 1
                    Else
                        i = i + 1
                        p = 0
                        ReDim Preserve Mbist(sheet_type).Block(counter).block_type_pat(BlockCounter)
                        ReDim Preserve Mbist(sheet_type).Block(counter).block_type_pat(BlockCounter).pat(p)
                        Mbist(sheet_type).Block(counter).block_type_pat(BlockCounter).pat(p) = myCell.offset(rowOffset:=0, columnOffset:=1)
                        pre_block_type = m_cellStr
                        
                        p = p + 1
                    End If
                    'p = p + 1
                Loop
                   p = p + 1
'''                ReDim Preserve block_name(counter - 1)
'''                ReDim Preserve block_name_count(counter - 1)
                ReDim Preserve block_name(counter01 - 1)
                ReDim Preserve block_name_count(counter01 - 1)
                m_Block_Cnt = n     '//Block amount
                m_Block_cata_Cnt = counter01
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================Pattern_Block
             Case idx_Pattern_Block_column
                pre_Character = vbNullString:: Character = vbNullString
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    pattern_name(n) = m_cellStr
                    n = n + 1
                Loop
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ReDim Preserve pattern_name(n - 1)
                m_Pat_Cnt = n
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================BinFlag Decision Type(TestInstanceName, GroupName)
             Case idx_BinFlag_Decision_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    inst_group_name(n) = m_cellStr
                    n = n + 1
                Loop
                ReDim Preserve inst_group_name(n - 1)
                m_flag_range = n
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================Match Name for BinFlag
             Case idx_Name_for_BinFlag_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    flag_match_name(n) = m_cellStr
                    n = n + 1
                Loop
                ReDim Preserve flag_match_name(n - 1)
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================BinFlag Mid Name
             Case idx_BinFlag_Mid_Name_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    flag_mid_name(n) = m_cellStr
                    n = n + 1
                Loop
                ReDim Preserve flag_mid_name(n - 1)
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================BinFlag with PM/BM (Yes/No)
             Case idx_BinFlag_with_PM_BM_column
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    If (UCase(m_cellStr) Like UCase("Yes")) Then
                        flag_pm_bm(n) = True
                    Else
                        flag_pm_bm(n) = False
                    End If

                    n = n + 1
                Loop
                ReDim Preserve flag_pm_bm(n - 1)
                n = 0:: counter = 0:: i = 0:: p = 0
            '=================================================================================TimingLevel
             Case idx_DCSpec_column
                Do While (n < m_lastNCnt) 'row count
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    Call Separate_nu_char(CStr(MySheet.Cells(n + 2, 5).value), number, Character)
                    
                    If (m_cellStr <> "") Then
                        Mbist(sheet_type).Block(counter).timing_levels.dc_spec = m_cellStr
                    End If
                    
                    If m_cellStr2 <> UCase(Character) Then
                        m_cellStr2 = Character
                        counter = counter + 1
                    End If
                    n = n + 1
                Loop
                n = 0:: counter = 0:: m_cellStr2 = vbNullString
             Case idx_DCSel_column
                Do While (n < m_lastNCnt) 'row count
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    Call Separate_nu_char(CStr(MySheet.Cells(n + 2, 5).value), number, Character)
                    
                    If (m_cellStr <> "") Then
                        Mbist(sheet_type).Block(counter).timing_levels.dc_sel = m_cellStr
                    End If
                    
                    If m_cellStr2 <> UCase(Character) Then
                        m_cellStr2 = Character
                        counter = counter + 1
                    End If
                    n = n + 1
                Loop
                n = 0:: counter = 0:: m_cellStr2 = vbNullString
             Case idx_ACSpec_column
                Do While (n < m_lastNCnt) 'row count
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    Call Separate_nu_char(CStr(MySheet.Cells(n + 2, 5).value), number, Character)
                    
                    If (m_cellStr <> "") Then
                        Mbist(sheet_type).Block(counter).timing_levels.ac_spec = m_cellStr
                    End If
                    
                    If m_cellStr2 <> UCase(Character) Then
                        m_cellStr2 = Character
                        counter = counter + 1
                    End If
                    n = n + 1
                Loop
                n = 0:: counter = 0:: m_cellStr2 = vbNullString
             Case idx_ACSel_column
                Do While (n < m_lastNCnt) 'row count
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    Call Separate_nu_char(CStr(MySheet.Cells(n + 2, 5).value), number, Character)
                    
                    If (m_cellStr <> "") Then
                        Mbist(sheet_type).Block(counter).timing_levels.ac_sel = m_cellStr
                    End If
                    
                    If m_cellStr2 <> UCase(Character) Then
                        m_cellStr2 = Character
                        counter = counter + 1
                    End If
                    n = n + 1
                Loop
                n = 0:: counter = 0:: m_cellStr2 = vbNullString
             Case idx_TimingSheet_column
                Do While (n < m_lastNCnt) 'row count
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    Call Separate_nu_char(CStr(MySheet.Cells(n + 2, 5).value), number, Character)
                    
                    If (m_cellStr <> "") Then
                        Mbist(sheet_type).Block(counter).timing_levels.timingsheet = m_cellStr
                    End If
                    
                    If m_cellStr2 <> UCase(Character) Then
                        m_cellStr2 = Character
                        counter = counter + 1
                    End If
                    n = n + 1
                Loop
                n = 0:: counter = 0:: m_cellStr2 = vbNullString
             Case idx_LevelSheet_column
                Do While (n < m_lastNCnt) 'row count
                    Set myCell = myCell.offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.value))
                    Call Separate_nu_char(CStr(MySheet.Cells(n + 2, 5).value), number, Character)
                    
                    If (m_cellStr <> "") Then
                        Mbist(sheet_type).Block(counter).timing_levels.levelsheet = m_cellStr
                    End If
                    
                    If m_cellStr2 <> UCase(Character) Then
                        m_cellStr2 = Character
                        counter = counter + 1
                    End If
                    n = n + 1
                Loop
                n = 0:: counter = 0:: m_cellStr2 = vbNullString
             Case Else
                'DebugPrintLog "6...Empty Column(" & M & ") !!!"
             End Select
            '=================================================================================

        Loop ''''end of Do While (m <= idx_END)
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++for BM & instance & pattern
         i = 0:: n = 0:: p = 0:: g = 0
         ReDim Preserve Mbist(sheet_type).Block(m_Block_cata_Cnt - 1)
         For i = 0 To m_Block_cata_Cnt - 1
            Mbist(sheet_type).Block(i).block_name = block_name(i)
            Mbist(sheet_type).Block(i).block_count = block_name_count(i)
            '--------------------------------------------------------------
            ReDim Preserve Mbist(sheet_type).Block(i).block_count_name(block_name_count(i) - 1)
            For n = 0 To block_name_count(i) - 1
                Mbist(sheet_type).Block(i).block_count_name(n) = block_count_name(i, n)
            Next n
            '--------------------------------------------------------------
            'Mbist(sheet_type).Block(i).ins_name = ins_true_name(i)
            '--------------------------------------------------------------
            p = 0
            For n = 0 To m_Pat_Cnt - 1
                'For counter = 0 To block_name_count(i) - 1
                    'If (UCase(pattern_name(n)) Like UCase("*" + Mbist(sheet_type).block(i).block_count_name(counter)) + "*") Then
                    Dim Block() As String
                    Block = Split(Mbist(sheet_type).Block(i).block_name, "_")
                    If UCase(pattern_name(n)) Like UCase("*" & Block(0) & "*") Then
                        ReDim Preserve Mbist(sheet_type).Block(i).pat_name(p)
                        ReDim Preserve Mbist(sheet_type).Block(i).pat_tested(p)
                        Mbist(sheet_type).Block(i).pat_name(p) = pattern_name(n)
                        Mbist(sheet_type).Block(i).pat_tested(p) = False
                        p = p + 1
                        Mbist(sheet_type).Block(i).pat_count = p
                        'Exit For
                    End If
                'Next counter
            Next n
            '--------------------------------------------------------------
            For p = 0 To UBound(Mbist(sheet_type).Block(i).block_type_pat)
                For g = 0 To UBound(Mbist(sheet_type).Block(i).block_type_pat(p).pat)
                    pre_store_nu = 0
                    ReDim Preserve Mbist(sheet_type).Block(i).block_type_pat(p).instance(g)
                    For n = 0 To UBound(ins_true_name)
                        ins_true_name(n) = Trim(ins_true_name(n))
                        Mbist(sheet_type).Block(i).block_type_pat(p).pat(g) = Trim(Mbist(sheet_type).Block(i).block_type_pat(p).pat(g))
                        Loc_dash = InStr(1, ins_true_name(n), Mbist(sheet_type).Block(i).block_type_pat(p).pat(g))
                        If (Loc_dash > 0) Then
                            If (pre_store_nu < Len(Mbist(sheet_type).Block(i).block_type_pat(p).pat(g))) Then
                                store_string = ins_true_name(n)
                                Mbist(sheet_type).Block(i).block_type_pat(p).instance(g) = ins_true_name(n)
                            End If
                            pre_store_nu = Len(Mbist(sheet_type).Block(i).block_type_pat(p).pat(g))
                        End If
                    Next n
                Next g
            Next p
            
            For p = 0 To UBound(Mbist(sheet_type).Block(i).block_type_pat)
              For g = 0 To UBound(Mbist(sheet_type).Block(i).block_type_pat(p).pat)
                 If (p > 0) Then
                    ReDim Preserve Mbist(sheet_type).Block(i).block_type_pat(p).instance(g)
                    Mbist(sheet_type).Block(i).block_type_pat(p).instance(g) = Mbist(sheet_type).Block(i).block_type_pat(0).instance(g)
                 End If
              Next g
            Next p
            '--------------------------------------------------------------
         Next i

         ReDim Preserve Mbist(sheet_type).Block(i - 1)
         i = 0:: n = 0:: p = 0
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++for flag
         For i = 0 To m_flag_range - 1
            If (inst_group_name(i) = UCase("TestInstanceName") And flag_pm_bm(i) = True) Then
                ReDim Preserve mbist_match(sheet_type).inst_nu(p)
                mbist_match(sheet_type).inst_nu(p).binflag_match_name = flag_match_name(i)
                mbist_match(sheet_type).inst_nu(p).binflag_mid_name = flag_mid_name(i)
                p = p + 1
            End If
         Next i
         mbist_match(sheet_type).inst_count = p
         i = 0:: n = 0:: p = 0
         '//flag=front_pp+performance+NV+Blockname
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    End If ''''end of If (find_AllHeader) Then
    
'    If Mbist(sheet_type).Block(0).block_type_pat(1).instance(0) Is Nothing Then
'       i = 0:: n = 0:: p = 0
'    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Separate_nu_char(word As String, ByRef number As Long, ByRef Character As String)

    Dim S As String
    Dim i As String
    Dim j As String
    Dim k As Long
    Dim divide_placement As Long

    For k = Len(word) To 1 Step -1
        If IsNumeric(mid(word, k, 1)) = False Then
            divide_placement = k
            Exit For
        End If
    Next k
    
''    For k = 1 To Len(word)
''        If IsNumeric(Mid(word, k, 1)) = True Then
''            j = j & Mid(word, k, 1)
''        Else
''            i = i & Mid(word, k, 1)
''        End If
''    Next k

    For k = 1 To Len(word)
        If k <= divide_placement Then
            i = i & mid(word, k, 1)
        Else
            j = j & mid(word, k, 1)
        End If
    Next k
    
    Character = i
    number = CLng(j)
    'If UCase(Character) = "APK" Then
    ''    number = CLng(j) - 1
    'Else
    ''    number = CLng(j)
    'End If
    
End Function


