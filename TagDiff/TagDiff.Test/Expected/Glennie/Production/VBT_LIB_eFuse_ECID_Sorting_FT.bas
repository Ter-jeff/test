Attribute VB_Name = "VBT_LIB_eFuse_ECID_Sorting_FT"
Option Explicit

''''''''''''''''''''''''''''''''''''''
Private Const VBT_LIB_eFuse_ECID_Sorting_FT = "VBT_LIB_eFuse_ECID_Sorting_FT"
Private Type ECID_DataStructure_S
    Fab_LotID As String
    WaferID As String
    Corr_X As String
    Corr_Y As String
    FT_HW_Bin As Long
    OtherBin As Long
    FT_ERR_Bin As Long
    
End Type

Public Type ECID_Struc_S
    Category() As ECID_DataStructure_S
End Type

Public ECID_CheckList_S As ECID_Struc_S
Public ECID_Dict_S As Object
Public Flag_ECID_Read As Boolean

' [20231113][All][CP] Add file type gating for ECID_Check_List_XXXX
' [20231113][All][CP] For OSAT ECID sorting, search ECID check list file first at the same path as the program
Public Function ECID_Dict_Buildup_S()
On Error GoTo errHandler
Dim funcName As String: funcName = "ECID_Dict_Buildup_S"
'Get the file path of ECID_CheckList_S
Dim pathstr As String
Dim Alpath As String
Dim url As String
Dim fd As String
Dim fs, f As Object
Dim line As String
Dim tempAL() As String
Dim wb As Workbook
Dim file_path_temp As String
Dim DataString As String
Dim wsh As Object
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim Flag_ECID_CheckList As Boolean
Dim my_wbook As Workbook
Dim my_sheet As Worksheet
Dim tempstr1 As String
Dim temp() As String
Dim ECID_Info_ary() As String
Dim cnt As Integer, Startidx As Long
Dim ECID_Info_idx As Integer 'Dim ECID_Info_idx As Intrger
Dim Lot_idx As Long
Dim Wafer_num_idx As Long
Dim X_idx As Long
Dim Y_idx As Long
Dim HB_idx  As Long
Dim Flag_LotID_Read As Boolean
Dim Flag_WaferID_Read As Boolean
Dim Flag_X_Read As Boolean
Dim Flag_Y_Read As Boolean
Dim Flag_HB_Read As Boolean
Dim q As Long
Dim tempInfo As Long

    Set wsh = VBA.CreateObject("WScript.Shell")
    '''Init flags'''
    Flag_ECID_Read = False
    Flag_ECID_CheckList = False

    '''-----------------------------------------------------------------------------------------------'''
    '''20231107 For ECID sorting in OSAT, search ECID check list file at the same path as the program '''
    '''If ECID check list file don't exist in the path, will get the its path from Autoload           '''
    '''-----------------------------------------------------------------------------------------------'''
    file_path_temp = TheExec.TestProgram.Path
    pathstr = file_path_temp & "\ECID_Check_List_XXXX.csv"

    If Dir(pathstr) = "" Then
        'wsh.Run "net use U: \\10.77.128.31\nt_c651eng\C651  C651ENG_user123 /user:10.77.128.31\NT_C651ENG_user", windowStyle, waitOnReturn
        url = "C:\Flex\Applications\"
        Set fs = CreateObject("Scripting.FileSystemObject")
    
        '''Check whether Autoload exists'''
        fd = Dir(url & "*.stp")
        If fd <> "" Then
            TheExec.Datalog.WriteComment "Autoload exists: " & fd
        Else
            Call Print_Error_Message(Error_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "Autoload doesn't exist in " & url & " !!")
            Exit Function
        End If
        Set f = fs.OpenTextFile(url & fd)
    
        '''Check whether ECID Check List path exists in Autoload'''
        Do Until f.AtEndOfStream
            line = f.ReadLine
            If UCase(line) Like UCase("ECID_file*") Then
                tempAL = Split(line, vbTab)
                pathstr = tempAL(2)
                Flag_ECID_CheckList = True
            End If
        Loop
    Else
        Flag_ECID_CheckList = True
    End If
    If True = Flag_ECID_CheckList And Dir(pathstr) <> "" Then
        TheExec.Datalog.WriteComment "From Autolad, ECID Check List exists: " & pathstr
    Else
        Call Print_Error_Message(Error_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "Autoload content doesn't have ECID_file path or ECID Check List doesn't exist in " & pathstr & " !!")
        Exit Function
    End If

    Set ECID_Dict_S = CreateObject("Scripting.Dictionary")

    If pathstr Like "*.csv" Then
    Else
        Call Print_Error_Message(Error_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "ECID Check List File type is NOT .csv!")
        Exit Function
    End If
    '''-------Parsing Table Start----------'''
    Open pathstr For Input As #2
        tempstr1 = Input(LOF(2), #2)
        temp = Split(tempstr1, Chr(10))
    Close #2

    '''Init flags'''
    Flag_LotID_Read = False
    Flag_WaferID_Read = False
    Flag_X_Read = False
    Flag_Y_Read = False
    Flag_HB_Read = False

    For cnt = 0 To UBound(temp)
        If UCase(temp(cnt)) Like UCase("*Fab*Lot*ID*") Then
            ECID_Info_ary = Split(temp(cnt), ",")
            For ECID_Info_idx = 0 To UBound(ECID_Info_ary)
                If UCase(Trim(ECID_Info_ary(ECID_Info_idx))) Like UCase("*Fab*Lot*ID*") Then
                    Lot_idx = ECID_Info_idx
                    Flag_LotID_Read = True
                ElseIf UCase(Trim(ECID_Info_ary(ECID_Info_idx))) Like UCase("*Wafer*ID*") Then
                    Wafer_num_idx = ECID_Info_idx
                    Flag_WaferID_Read = True
                ElseIf UCase(Trim(ECID_Info_ary(ECID_Info_idx))) = UCase("X") Then
                    X_idx = ECID_Info_idx
                    Flag_X_Read = True
                ElseIf UCase(Trim(ECID_Info_ary(ECID_Info_idx))) = UCase("Y") Then
                    Y_idx = ECID_Info_idx
                    Flag_Y_Read = True
                ElseIf UCase(Trim(ECID_Info_ary(ECID_Info_idx))) Like UCase("*HW*Bin*") Then
                    HB_idx = ECID_Info_idx
                    Flag_HB_Read = True
                Else
                End If
            Next ECID_Info_idx
            Startidx = cnt + 1
            Exit For
        End If
    Next cnt

    '''Check whether ECID Check List header is correct'''
    If False = Flag_LotID_Read Or False = Flag_WaferID_Read Or False = Flag_X_Read Or False = Flag_Y_Read Or False = Flag_HB_Read Then
        Call Print_Error_Message(Error_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "ECID Check List Header has issue! Please check Header in ECID Check List!!")
        Exit Function
    End If

    q = 0
    ReDim Preserve ECID_CheckList_S.Category(q)
    For tempInfo = Startidx To UBound(temp)
        If UCase(Trim(temp(tempInfo))) Like UCase("*END*") Then
            Exit For
        End If
        '''Check whether ECID Check List has empty value'''
        ECID_Info_ary = Split(temp(tempInfo), ",")
        If vbNullString = ECID_Info_ary(Lot_idx) Or vbNullString = ECID_Info_ary(Wafer_num_idx) Or vbNullString = ECID_Info_ary(X_idx) Or vbNullString = ECID_Info_ary(Y_idx) Or vbNullString = ECID_Info_ary(HB_idx) Then
            Call Print_Error_Message(Error_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "There is empty value in ECID Check List!!")
            Exit Function
        End If
        With ECID_CheckList_S.Category(q)
            .Fab_LotID = ECID_Info_ary(Lot_idx)
            .WaferID = ECID_Info_ary(Wafer_num_idx)
            .Corr_X = ECID_Info_ary(X_idx)
            .Corr_Y = ECID_Info_ary(Y_idx)
            .FT_HW_Bin = ECID_Info_ary(HB_idx)
            
            '''Create DataString for storing into ECID dictionary'''
            DataString = .Fab_LotID & "," & .WaferID & "," & .Corr_X & "," & .Corr_Y
             
            '''Make sure all data are not duplicate, otherwise it would occur error!!'''
            If Not ECID_Dict_S.Exists(DataString) Then
                ECID_Dict_S.Add DataString, .FT_HW_Bin
            Else
            End If
        End With
    Next tempInfo
    '''-------Parsing Table End----------'''
    Flag_ECID_Read = True

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "Something error when creating ECID DataBase!!")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function BinSorting_Compare_FT_ECID_S()
On Error GoTo errHandler
Dim funcName As String: funcName = "BinSorting_Compare_FT_ECID_S"
Dim CurTime As String
Dim key As New SiteVariant
Dim Curpath As String
Dim RecordDate As String
Dim SiteCnt As Integer
Dim site As Variant
Dim Site_Idx As Long
Dim Site_select() As Long
Dim ECID_Found_Flag As New SiteBoolean
Dim m_loop As Long
Dim Comp_Result As New SiteLong
Dim n As Integer

    '''---------------------------------------------------'''
    '''Check whether parsing ECID check list is Successful'''
    '''If it fails, parse ECID check list again           '''
    '''---------------------------------------------------'''
    If Flag_ECID_Read = False Then
        Call Print_Error_Message(Warning_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "ECID check list doesn't exist!!")
        TheExec.Datalog.WriteComment "Parse ECID check list again!"
        Call ECID_Dict_Buildup_S
        If Flag_ECID_Read = False Then
            Exit Function
        Else
            TheExec.Datalog.WriteComment "Parsing ECID check list is Successful!"
        End If
    End If
    '''----------------------------------------------------------------------------------------------------'''
    '''Non-2C Mode                                                                                         '''
    '''If the die exists, the result will pass and bin number will be that user can set in ECID check list.'''
    '''----------------------------------------------------------------------------------------------------'''
    If False = EFUSE_ECID_SORTING_2CMODE Then
        For Each site In TheExec.sites
            key(site) = HramLotId(site) & "," & HramWaferId(site) & "," & HramXCoord(site) & "," & HramYCoord(site)
            If ECID_Dict_S.Exists(key(site)) Then
                TheExec.sites(site).BinNumber = ECID_Dict_S(key(site))
                TheExec.sites(site).SortNumber = ECID_Dict_S(key(site))
                TheExec.flow.TestLimit 1, 0, 1, , , , , , "BinSorting_pass"
            Else
                TheExec.sites(site).BinNumber = 1
                TheExec.sites(site).SortNumber = 1
                TheExec.flow.TestLimit 10, 0, 1, , , , , , "BinSorting_fail"
                Call Print_Error_Message(Warning_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "ECID is NOT in excel sheet!!" & vbCrLf)
            End If
        Next site
    '''---------------------------------------------------------------------'''
    '''2C Mode                                                              '''
    '''The result is pass, When Master or Slave die exit in ECID Check List.'''
    '''---------------------------------------------------------------------'''
    Else
        TheExec.Datalog.WriteComment "< Specific ECID Sorting >"
        'For All Project'
        SiteCnt = TheExec.sites.Existing.Count
        Site_Idx = SiteCnt / 2
        ReDim Site_select(Site_Idx) As Long
        For n = 0 To Site_Idx - 1
            Site_select(n) = n
        Next n

        For m_loop = 0 To UBound(Site_select) - 1
            For Each site In TheExec.sites
                If CLng(site) = Site_select(m_loop) Then  ')site = 0 Or site = 1 Then
                    ECID_Found_Flag(site) = False
                    ECID_Found_Flag(site + Site_Idx) = False
                    key(site) = HramLotId(site) & "," & HramWaferId(site) & "," & HramXCoord(site) & "," & HramYCoord(site)
                    key(site + Site_Idx) = HramLotId((site + Site_Idx)) & "," & HramWaferId((site + Site_Idx)) & "," & HramXCoord((site + Site_Idx)) & "," & HramYCoord((site + Site_Idx))
                    If ECID_Dict_S.Exists(key(site)) Then
                        ECID_Found_Flag(site) = True
                        TheExec.sites(site).BinNumber = ECID_Dict_S(key(site))
                        TheExec.sites(site).SortNumber = ECID_Dict_S(key(site))
                        TheExec.sites(site + Site_Idx).BinNumber = ECID_Dict_S(key(site))
                        TheExec.sites(site + Site_Idx).SortNumber = ECID_Dict_S(key(site))
                        TheExec.flow.TestLimit 1, 0, 1, , , , , , "BinSorting_pass"
                        Comp_Result(site) = 1
                        Comp_Result(site + Site_Idx) = 1
                    End If
                    If ECID_Dict_S.Exists(key(site + Site_Idx)) And ECID_Found_Flag(site) = False Then
                        ECID_Found_Flag(site + Site_Idx) = True
                        TheExec.sites(site).BinNumber = ECID_Dict_S(key(site + Site_Idx))
                        TheExec.sites(site).SortNumber = ECID_Dict_S(key(site + Site_Idx))
                        TheExec.sites(site + Site_Idx).BinNumber = ECID_Dict_S(key(site + Site_Idx))
                        TheExec.sites(site + Site_Idx).SortNumber = ECID_Dict_S(key(site + Site_Idx))
                        Comp_Result(site) = 1
                        Comp_Result(site + Site_Idx) = 1
                        TheExec.flow.TestLimit 1, 0, 1, , , , , , "BinSorting_pass"
                    End If
                    If ECID_Found_Flag(site) = False And ECID_Found_Flag(site + Site_Idx) = False Then
                        TheExec.sites(site).BinNumber = 1
                        TheExec.sites(site).SortNumber = 1
                        TheExec.sites(site + Site_Idx).BinNumber = 1
                        TheExec.sites(site + Site_Idx).SortNumber = 1
                        Comp_Result(site) = 0
                        Comp_Result(site + Site_Idx) = 0
                        TheExec.flow.TestLimit 10, 0, 1, , , , , , "BinSorting_fail"
                        Call Print_Error_Message(Warning_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "ECID is NOT in excel sheet!!" & vbCrLf)
                    End If
                End If
            Next site
        Next m_loop
        TheExec.flow.TestLimit resultVal:=Comp_Result, lowVal:=1, hiVal:=1, Tname:="ECID Sorting", ForceResults:=tlForceNone
    End If
Exit Function
errHandler:
     Call Print_Error_Message(Error_Info, VBT_LIB_eFuse_ECID_Sorting_FT, funcName, "Please check it out.!!")
     If AbortTest Then Exit Function Else Resume Next
End Function
