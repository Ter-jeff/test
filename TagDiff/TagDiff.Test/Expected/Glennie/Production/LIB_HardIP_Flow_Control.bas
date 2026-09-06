Attribute VB_Name = "LIB_HardIP_Flow_Control"

Public Function HardIP_ReduceMemory_ALL_TTR()

    Call Disable_HIP_MN_And_CZ
    Call Move_PROD_EnableWord_to_JOB_forMemoryTTR
    Call Disable_NoUse_PP

End Function



Public Function Check_Irange_for_TTR()

    Dim I_range_CWCW As Scripting.Dictionary

    Set I_range_CWCW = CreateObject("Scripting.Dictionary")

    'range("A5").EntireRow.Insert
    Dim x As Long
    Dim y As Long
    
    Dim c7
    Dim c10
    Dim c20
    Dim Irange_arr() As String
    Dim Irange_arr_second() As String
    Dim xx As Variant
    Dim xxx As Variant
    Dim check_flag As Boolean: check_flag = False
    
    c7 = range("AE5:AE" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
    'c10 = range("H5:H" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value

    'ReDim c20(1 To UBound(c10), 1 To 1)
    
    For x = 1 To UBound(c7)
        Irange_arr = Split(c7(x, 1), "+")
        
        For Each xx In Irange_arr
            Irange_arr_second = Split(xx, ",")
            For Each xxx In Irange_arr_second
                I_range_CWCW(xxx) = 0
            Next xxx
            If InStr(xx, ",") <> 0 And I_range_CWCW.Count = 1 Then
                check_flag = True
            End If
            I_range_CWCW.RemoveAll
        Next xx
        
        If check_flag = True Then Sheets(ActiveSheet.name).Cells(x + 4, "AE").Interior.ColorIndex = 26
        check_flag = False
    Next x
    
    'range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c20
    'range("A5").EntireRow.delete

End Function




Public Function Only_Enable_HardIP_T1_Trim()

    range("A5").EntireRow.Insert
    Dim x As Long
    Dim y As Long
    
    Dim c7
    Dim c10
    Dim c20
    
    Dim str(9) As String, str_all As String
    
    str(0) = "CIO3PLL_CIO3PLLT1_PP_STNA0_C_FULP_AN_AA10_DLL_JTG_VIX_ALLFRV_SI_CIO3PLL_T1_NV"
    str(1) = "CIOAUSPLL_CAUSPLLT1_PP_STNA0_C_FULP_AN_AA24_DLL_JTG_CAL_ALLFRV_SI_CAUSPLL_T1_NV"
    str(2) = "DPTXAUSPLL_DAUSPLLT1_PP_STNA0_C_FULP_AN_AA23_DLL_JTG_CAL_ALLFRV_SI_DAUSPLL_T1_NV"
    str(3) = "LPDDR5PLL_LP5FCALT1_PP_STNA0_C_FULP_AN_AAL5_MEA_JTG_LP5_ALLFRV_SI_LP5FCAL_T1_NV"
    str(4) = "LPDPAUSPLL_LAUSPLLT1_PP_STNA0_C_FULP_AN_AA23_DLL_JTG_CAL_ALLFRV_SI_LAUSPLL_T1_NV"
    str(5) = "PCIEAUSPLL_PAUSPLLT1_PP_STNA0_C_FULP_AN_AA25_DLL_JTG_CAL_ALLFRV_SI_PAUSPLL_T1_NV"
    str(6) = "PCIEREFPLL_PCIERPLLT1_PP_STNA0_C_FULP_AN_AA21_DLL_JTG_CAL_ALLFRV_SI_PCIERPLL_T1_NV"
    str(7) = "PCIETXPLL_PCIEPLLT1_PP_STNA0_C_FULP_AN_AA22_DLL_JTG_CAL_ALLFRV_SI_PCIETPLL_T1_NV"
    str(8) = "CIO_VREGCM3_PP_STNA0_S_FULP_AN_CICM_MEA_JTG_CAL_ALLFRV_SI_VREG_CM3_NV"
    str(9) = "PCIE_ITRIMCM3_PP_STNA0_S_FULP_AN_PCCM_MEA_JTG_CAL_ALLFRV_SI_VREG_CM3_NV"

    str_all = Join(str, ",")
    
    c7 = range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
    c10 = range("H5:H" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value

    ReDim c20(1 To UBound(c10), 1 To 1)
    
    For x = 1 To UBound(c10)
        If UCase(c7(x, 1)) = "TEST" And InStr(str_all, c10(x, 1)) <> 0 Then
            c20(x, 1) = c7(x, 1)
        ElseIf UCase(c7(x, 1)) = "TEST" And (InStr(c10(x, 1), "Header") <> 0 Or InStr(c10(x, 1), "Footer") <> 0) Then
            c20(x, 1) = c7(x, 1)
        ElseIf UCase(c7(x, 1)) = "TEST" Or LCase(c7(x, 1)) = "characterize" Then
            c20(x, 1) = "nop"
        Else
            c20(x, 1) = c7(x, 1)
        End If
    Next x
    
    range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c20
    range("A5").EntireRow.delete

End Function





Public Function HIP_all_sheet(ByRef sheets_str_arr() As String)



Dim i As Long
Dim j As Long
Dim Arr_all_HIP() As String

For i = 1 To Worksheets.Count
    If Sheets(i).name Like "Flow_HARDIP_*" And (InStr(UCase(Sheets(i).name), "DRAM") = 0) Then
        j = j + 1
    End If
Next i
ReDim Arr_all_HIP(j - 1)
j = 0


For i = 1 To Worksheets.Count

    If Sheets(i).name Like "Flow_HARDIP_*" And (InStr(UCase(Sheets(i).name), "DRAM") = 0) Then
    
        Arr_all_HIP(j) = Sheets(i).name
        j = j + 1
        
    End If
Next i




'ReDim Arr_all_HIP(j - 1) As String

'Arr_all_HIP(0) = "Flow_HARDIP_CLKMON"
'Arr_all_HIP(1) = "Flow_HARDIP_ADCLK"
'Arr_all_HIP(2) = "Flow_HARDIP_AMPLP5"
'Arr_all_HIP(3) = "Flow_HARDIP_AMPLP5_PP_MN_HV"
'Arr_all_HIP(4) = "Flow_HARDIP_AMPLP5_PP_MN_LV"
'Arr_all_HIP(5) = "Flow_HARDIP_AMPLP5_PP_MN_NV"
'Arr_all_HIP(6) = "Flow_HARDIP_CIO"
'Arr_all_HIP(7) = "Flow_HARDIP_CIO3PLL"
'Arr_all_HIP(8) = "Flow_HARDIP_CIOAUSPLL"
'Arr_all_HIP(9) = "Flow_HARDIP_CIOREFPLL"
'Arr_all_HIP(10) = "Flow_HARDIP_DPTXAUSPLL"
'Arr_all_HIP(11) = "Flow_HARDIP_DSGRINGS"
'Arr_all_HIP(12) = "Flow_HARDIP_EUSB"
'Arr_all_HIP(13) = "Flow_HARDIP_FLPPLL"
'Arr_all_HIP(14) = "Flow_HARDIP_FRO"
'Arr_all_HIP(15) = "Flow_HARDIP_IO"
'Arr_all_HIP(16) = "Flow_HARDIP_JTAG"
'Arr_all_HIP(17) = "Flow_HARDIP_LPDDR5PLL"
'Arr_all_HIP(18) = "Flow_HARDIP_LPDPAUSPLL"
'Arr_all_HIP(19) = "Flow_HARDIP_LPDPRX"
'Arr_all_HIP(20) = "Flow_HARDIP_LPDPTX"
'Arr_all_HIP(21) = "Flow_HARDIP_MIPI"
'Arr_all_HIP(22) = "Flow_HARDIP_MISC"
'Arr_all_HIP(23) = "Flow_HARDIP_MTR"
'Arr_all_HIP(24) = "Flow_HARDIP_MTRBTS"
'Arr_all_HIP(25) = "Flow_HARDIP_MTRTDIODE"
'Arr_all_HIP(26) = "Flow_HARDIP_MTRTSNS"
'Arr_all_HIP(27) = "Flow_HARDIP_PCIE"
'Arr_all_HIP(28) = "Flow_HARDIP_PCIEAUSPLL"
'Arr_all_HIP(29) = "Flow_HARDIP_PCIEREFBUF"
'Arr_all_HIP(30) = "Flow_HARDIP_PCIETXPLL"
'Arr_all_HIP(31) = "Flow_HARDIP_RLXOSC"
'Arr_all_HIP(32) = "Flow_HARDIP_SEPVME"
'Arr_all_HIP(33) = "Flow_HARDIP_SOCPLL"
'Arr_all_HIP(34) = "Flow_HARDIP_UCSDM"
'Arr_all_HIP(35) = "Flow_HARDIP_ULPPLL"
'Arr_all_HIP(36) = "Flow_HARDIP_IVDM"


sheets_str_arr = Arr_all_HIP


End Function

Public Function Enable_HIP_MN_CZ()


    Dim Enable_Type As Integer
    
    Enable_Type = 1
    ' 1 : EnableMN
    ' 2 : EnableCZ
    ' 3 : EnableMNandCZ

    Dim x As Long
    Dim y As Long
    
    Dim c7
    Dim c10
    Dim c11
    Dim c20
    Dim sheet_name As Variant
    Dim sheets_str_arr() As String
    
    Call HIP_all_sheet(sheets_str_arr)


    For Each sheet_name In sheets_str_arr
        Worksheets(sheet_name).Activate
        range("A5").EntireRow.Insert

        c7 = range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c10 = range("C5:C" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11 = range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value

        For x = 2 To UBound(c7)
            If UCase(c7(x, 1)) = "TEST" Or UCase(c7(x, 1)) = "CALL" Or LCase(c7(x, 1)) = "characterize" Then
                If c11(x, 1) Like "*Memory*" And Enable_Type = 3 Then
                    c11(x, 1) = vbNullString
                ElseIf c11(x, 1) = "MemoryMN" And Enable_Type = 1 Then
                    c11(x, 1) = vbNullString
                ElseIf c11(x, 1) = "MemoryCZ" And Enable_Type = 2 Then
                    c11(x, 1) = vbNullString
                End If
            End If
        Next x
        
        range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11
        range("A5").EntireRow.delete
        Debug.Print sheet_name & "...done"
        
    Next sheet_name
    Debug.Print "Finish!!!"
End Function



Public Function Disable_HIP_MN_And_CZ()


    Dim x As Long
    Dim y As Long
    
    Dim c7
    Dim c10
    Dim c11
    Dim c11_bk
    Dim c20
    Dim sheet_name As Variant
    Dim sheets_str_arr() As String
    
    Dim cell As range
    
    
    
    Call HIP_all_sheet(sheets_str_arr)

    For Each sheet_name In sheets_str_arr
        Worksheets(sheet_name).Activate
        range("A5").EntireRow.Insert

        c7 = range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c10 = range("C5:C" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11 = range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11_bk = range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
    
        
        For x = 2 To UBound(c7)
            If UCase(c7(x, 1)) = "TEST" Or UCase(c7(x, 1)) = "CALL" Or LCase(c7(x, 1)) = "characterize" Then
                If InStr(UCase(c10(x, 1)), "HARDIP_NV_MN") > 0 Or InStr(UCase(c10(x, 1)), "HARDIP_LV_MN") > 0 Or InStr(UCase(c10(x, 1)), "HARDIP_HV_MN") > 0 Or InStr(UCase(c10(x, 1)), "HARDIP_NV_CZ") > 0 Or InStr(UCase(c10(x, 1)), "HARDIP_LV_CZ") > 0 Or InStr(UCase(c10(x, 1)), "HARDIP_HV_CZ") > 0 Then
                    If InStr(UCase(c10(x, 1)), "PROD_") = 0 Then

                        If UCase(c11(x, 1)) = "" Then
                                
                            If UCase(c10(x, 1)) Like "*MN*" Then
                                c11_bk(x, 1) = c11(x, 1)
                                c11(x, 1) = "MemoryMN"
                                If c11_bk(x, 1) <> "" Then
                                    c11_bk(x, 1) = c11_bk & "_MemoryBK"
                                Else
                                    c11_bk(x, 1) = "_MemoryBK"
                                End If
    
                            ElseIf UCase(c10(x, 1)) Like "*CZ*" Then
                                c11_bk(x, 1) = c11(x, 1)
                                c11(x, 1) = "MemoryCZ"
                                If c11_bk(x, 1) <> "" Then
                                    c11_bk(x, 1) = c11_bk & "_MemoryBK"
                                Else
                                    c11_bk(x, 1) = "_MemoryBK"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next x
        
        range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11
        range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11_bk
        range("A5").EntireRow.delete
        Debug.Print sheet_name & "...done"
        
    Next sheet_name
    Debug.Print "Finish!!!"
End Function


Public Function Disable_HIP_ROLLBACK()


    Dim x As Long
    Dim y As Long
    
    Dim c7
    Dim c10
    Dim c11
    Dim c11_bk
    Dim c20
    Dim sheet_name As Variant
    Dim sheets_str_arr() As String
    

    
    
    
    Call HIP_all_sheet(sheets_str_arr)

    For Each sheet_name In sheets_str_arr
        Worksheets(sheet_name).Activate
        range("A5").EntireRow.Insert

        c7 = range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c10 = range("C5:C" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11 = range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11_bk = range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        For x = 2 To UBound(c7)
            If UCase(c7(x, 1)) = "TEST" Or UCase(c7(x, 1)) = "CALL" Or LCase(c7(x, 1)) = "characterize" Then
                If c11_bk(x, 1) Like "_MemoryBK" Then
                    c11(x, 1) = Replace(c11_bk(x, 1), "_MemoryBK", "")
                    c11_bk(x, 1) = ""
                End If
            End If
            


        Next x
        
        range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11
        range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11_bk
        range("A5").EntireRow.delete
        Debug.Print sheet_name & "...done"
        
    Next sheet_name
    Debug.Print "Finish!!!"
End Function



Public Function Disable_HIP_ROLLBACK_T0TX()


    Dim x As Long
    Dim y As Long
    
    Dim c7
    Dim c10
    Dim c11
    Dim c11_bk
    Dim c20
    Dim sheet_name As Variant
    Dim sheets_str_arr() As String
    

    
    
    
    Call HIP_all_sheet(sheets_str_arr)

    For Each sheet_name In sheets_str_arr
        Worksheets(sheet_name).Activate
        range("A5").EntireRow.Insert

        c7 = range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c10 = range("C5:C" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11 = range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11_bk = range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        For x = 2 To UBound(c7)
            If UCase(c7(x, 1)) = "TEST" Or UCase(c7(x, 1)) = "CALL" Or LCase(c7(x, 1)) = "characterize" Then
                If c11_bk(x, 1) Like "_MemoryBK" Then
                    If c11(x, 1) Like "Memory*" Then
                    
                    Else
                
                        c11(x, 1) = Replace(c11_bk(x, 1), "_MemoryBK", "")
                        c11_bk(x, 1) = ""
                    End If
                    
                End If
            End If
            


        Next x
        
        range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11
        range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11_bk
        range("A5").EntireRow.delete
        Debug.Print sheet_name & "...done"
        
    Next sheet_name
    Debug.Print "Finish!!!"
End Function



Public Function Move_PROD_EnableWord_to_JOB_forMemoryTTR()


    Dim x As Long
    Dim y As Long
    
    Dim c7
    Dim c10
    Dim c11
    Dim c11_bk
    Dim c20
    Dim sheet_name As Variant
    Dim sheets_str_arr() As String
    Dim temp_str As String
    Call HIP_all_sheet(sheets_str_arr)

    For Each sheet_name In sheets_str_arr
        Worksheets(sheet_name).Activate
        range("A5").EntireRow.Insert

        c7 = range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c10 = range("C5:C" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11 = range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11_bk = range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        For x = 2 To UBound(c7)

            temp_str = vbNullString
            If UCase(c7(x, 1)) = "TEST" Or UCase(c7(x, 1)) = "CALL" Or LCase(c7(x, 1)) = "characterize" Then
                If UCase(c11(x, 1)) = "" Then
                    If UCase(c10(x, 1)) Like "*PROD_CP1*" Then
                        temp_str = temp_str & ",CP1"
                    End If
                    If UCase(c10(x, 1)) Like "*PROD_CP2*" Then
                        temp_str = temp_str & ",CP2"
                    End If
                    If UCase(c10(x, 1)) Like "*PROD_FT1*" Then
                        temp_str = temp_str & ",WLFT1,WLFT2,FT1"
                    End If
                    If UCase(c10(x, 1)) Like "*PROD_FT2*" Then
                        temp_str = temp_str & ",FT2"
                    End If
                    If UCase(c10(x, 1)) Like "*PROD_FT3*" Then
                        temp_str = temp_str & ",FT3"
                    End If
                    
                    If left(temp_str, 1) = "," Then temp_str = mid(temp_str, 2)
                    c11_bk(x, 1) = c11(x, 1)
                    c11(x, 1) = temp_str
                    If temp_str <> vbNullString Then
                        If c11_bk(x, 1) <> "" Then
                            c11_bk(x, 1) = c11_bk & "_MemoryBK"
                        Else
                            c11_bk(x, 1) = "_MemoryBK"
                        End If
                    End If
                End If
            End If
        Next x
        
        range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11
        range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11_bk
        range("A5").EntireRow.delete
        Debug.Print sheet_name & "...done"
        
    Next sheet_name
    Debug.Print "Finish!!!"
End Function


Public Function Disable_NoUse_PP()


    Dim x As Long
    Dim y As Long
    
    Dim c8
    Dim c7
    Dim c10
    Dim c11
    Dim c11_bk
    Dim c20
    Dim sheet_name As Variant
    Dim sheets_str_arr() As String
    
    Call HIP_all_sheet(sheets_str_arr)

    For Each sheet_name In sheets_str_arr
        Worksheets(sheet_name).Activate
        
        If InStr(sheet_name, "MTRB") <> 0 Then
            sheet_name = sheet_name
        End If
        
        
        range("A5").EntireRow.Insert

        c7 = range("G5:G" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c10 = range("C5:C" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11 = range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c8 = range("H5:H" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        c11_bk = range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value
        For x = 2 To UBound(c7)
            If UCase(c7(x, 1)) = "TEST" Or UCase(c7(x, 1)) = "CALL" Or LCase(c7(x, 1)) = "characterize" Then
                
                If InStr(LCase(c8(x, 1)), "fuse") = 0 And InStr(LCase(c8(x, 1)), "relay") = 0 And InStr(LCase(c8(x, 1)), "_read") = 0 And InStr(LCase(c8(x, 1)), "_prewrite") = 0 And InStr(LCase(c8(x, 1)), "_nwire") = 0 Then
                    If InStr(UCase(c10(x, 1)), "PROD") = 0 And c10(x, 1) <> "" Then
                        If UCase(c11(x, 1)) = "" Then
                            c11(x, 1) = "MemoryNoUsePP"
                            c11_bk(x, 1) = "_MemoryBK"
                        End If
                    End If
                End If
                
            

            End If
        Next x
        
'        For x = 2 To UBound(c7)
'            If c8(x, 1) Like "*Header*" Or c8(x, 1) Like "*Footer*" Then
'                    If c11(x, 1) = "MemoryNoUsePP" Then
'                        c11(x, 1) = ""
'                    End If
'            End If
'        Next x
        
        range("D5:D" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11
        range("A5:A" + CStr(Sheets(ActiveSheet.name).UsedRange.Rows.Count)).value = c11_bk
        range("A5").EntireRow.delete
        Debug.Print sheet_name & "...done"
        
    Next sheet_name
    Debug.Print "Finish!!!"
End Function
