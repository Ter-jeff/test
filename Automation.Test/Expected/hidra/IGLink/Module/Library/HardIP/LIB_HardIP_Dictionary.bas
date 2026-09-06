Attribute VB_Name = "LIB_HardIP_Dictionary"
Option Explicit
Public gDictDebug As New Dictionary                                                         'MTR Record
'**************************************************
'SeaHawk Edited by 20190606
Public gl_SpecialString As String
Public gl_DictString As New Dictionary
Public gl_DictDSPWave As New Dictionary
'**************************************************
Public gDictDSPWaves As New Dictionary
Private gDictCurrMeasurements As New Dictionary
Private gDictSiteLong As New Dictionary
Public gDictSiteDouble As New Dictionary   'Support Store SiteDouble -- 20230627
'Private RegDict As New Dictionary

' Function to retrieve a current/voltage/frequency measurement. This
' function can be called by user interpose functions to access
' previously stored measurements
Public Function GetStoredMeasurement(keyname As String) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    keyname = LCase(keyname)
    If Not gDictCurrMeasurements.Exists(keyname) Then
        TheExec.ErrorLogMessage "Stored measurement " & keyname & " not found."
    Else
        Set GetStoredMeasurement = gDictCurrMeasurements(keyname)
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "GetStoredMeasurement") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

' Function to store a measurement for later retrieval, typically from a custom user
' postMeasure interpose function
Public Function AddStoredMeasurement(keyname As String, ByVal obj As Variant)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim temp_obj As New PinListData
    temp_obj = obj
    keyname = LCase(keyname)
    If gDictCurrMeasurements.Exists(keyname) Then
        gDictCurrMeasurements.Remove (keyname)
    End If
    gDictCurrMeasurements.Add keyname, temp_obj
    Set temp_obj = Nothing

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "AddStoredMeasurement") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

' Function to retrieve a captured DSPWave. This
' function can be called by user interpose functions to access
' previously stored data
Public Function GetStoredCaptureData(keyname As String) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    keyname = LCase(keyname)
    If Not gDictDSPWaves.Exists(keyname) Then
        TheExec.ErrorLogMessage "Stored capture data " & keyname & " not found."
    Else
        Set GetStoredCaptureData = gDictDSPWaves(keyname)
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "GetStoredCaptureData") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

' Function to store a measurement for later retrieval, typically from a custom user
' postMeasure interpose function
Public Function AddStoredCaptureData(keyname As String, ByVal obj As DSPWave)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim temp_obj As New DSPWave
    temp_obj = obj
    keyname = LCase(keyname)
    If gDictDSPWaves.Exists(keyname) Then
        gDictDSPWaves.Remove (keyname)
    End If
    gDictDSPWaves.Add keyname, temp_obj
    
    '20220106, Add for Real value validation( reverse bit)
    Dim vsite As Variant
    Dim objvalue As New DSPWave
    If gB_efuse_DicValue_Chk_Flag = True Then ''Efuse_DicValue_Chk
        For Each vsite In TheExec.sites
            If temp_obj.SampleSize = 1 Then
                objvalue = temp_obj.ConvertDataTypeTo(DspLong)
                TheExec.Datalog.WriteComment "Site(" + CStr(vsite) + ")" + "@@@ Key_Name:= " + keyname + " Value:= " + CStr(objvalue.Element(0))
            Else
                objvalue = temp_obj.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, temp_obj.SampleSize, 0)
                TheExec.Datalog.WriteComment "Site(" + CStr(vsite) + ")" + "@@@ Key_Name:= " + keyname + " Value:= " + CStr(objvalue.Element(0))
            End If
        Next vsite
    End If
    ''Efuse_DicValue_Chk <==
    Set temp_obj = Nothing

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "AddStoredCaptureData") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function RemoveAllStored()
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    gDictCurrMeasurements.RemoveAll
    gDictDSPWaves.RemoveAll
    gDictSiteLong.RemoveAll
	gDictSiteDouble.RemoveAll
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "RemoveAllStored") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function AddStoredData(keyname As String, ByVal obj As SiteDouble)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim temp_obj As New SiteDouble
    temp_obj = obj
    keyname = LCase(keyname)
    If gDictSiteLong.Exists(keyname) Then
        gDictSiteLong.Remove (keyname)
    End If
    gDictSiteLong.Add keyname, temp_obj
    Set temp_obj = Nothing

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "AddStoredData") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function GetStoredData(keyname As String) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    keyname = LCase(keyname)
    If Not gDictSiteLong.Exists(keyname) Then
        TheExec.ErrorLogMessage "Stored capture data " & keyname & " not found."
    Else
        Set GetStoredData = gDictSiteLong(keyname)
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "GetStoredData") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function GetRegFromDictByTestByMode(RegAssignment As String, RegAssignChecker As Boolean) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

If RegDict.Exists(RegAssignment) Then
    RegAssignChecker = True
    RegAssignment = RegDict(RegAssignment)
Else
    RegAssignChecker = False
    'TheExec.ErrorLogMessage ("Your RegAssignment is not specified in Reg_Assign Sheet")
    'Debug.Print ("Your RegAssignment is not specified in Reg_Assign Sheet")
End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "GetRegFromDictByTestByMode") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Public_AddStoredString(keyname As String, ByRef obj As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    keyname = LCase(keyname)
    If gl_DictString.Exists(keyname) Then
        gl_DictString.Remove (keyname)
    End If
    gl_DictString.Add keyname, obj
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "Public_AddStoredString") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


''20190604AddFunction
Public Function IsExists_StoredCaptureData(keyname As String) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    keyname = LCase(keyname)
    IsExists_StoredCaptureData = gDictDSPWaves.Exists(keyname)
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "IsExists_StoredCaptureData") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Public_GetStoredString(keyname As String) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    keyname = LCase(keyname)
    If Not gl_DictString.Exists(keyname) Then
        TheExec.ErrorLogMessage "Stored measurement " & keyname & " not found."
    Else
'        Set Public_GetStoredString = gl_DictString(KeyName)
        gl_SpecialString = gl_DictString(keyname)
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "Public_GetStoredString") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function GetStoredString(keyname As String) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

'**************************************************
'LLPS CChop CZ 20200226
'**************************************************
    keyname = LCase(keyname)
    If Not gl_DictString.Exists(keyname) Then
        TheExec.ErrorLogMessage "Stored measurement " & keyname & " not found."
    Else
         GetStoredString = gl_DictString(keyname)
        'gl_SpecialString = gl_DictString(KeyName)
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "GetStoredString") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function StoreDataAllType(keyname As String, ByVal obj As Variant)
    'Store data to Dictionary for Multi-Data Type -- 20230627
    'DataType : SiteDouble/SiteLong/PinListData/DSPWave
    Dim SD_temp_obj As New SiteDouble
    Dim SL_temp_obj As New SiteLong
    Dim PLD_temp_obj As New PinListData
    Dim DSPWF_temp_obj As New DSPWave
    Dim Vsite As Variant
    Dim objvalue As New DSPWave
    keyname = LCase(keyname)
    If InStr(UCase(TypeName(obj)), "IPINLISTDATA") <> 0 Then
        PLD_temp_obj = obj
        If gDictCurrMeasurements.Exists(keyname) Then
            gDictCurrMeasurements.Remove (keyname)
        End If
        gDictCurrMeasurements.Add keyname, PLD_temp_obj
        Set PLD_temp_obj = Nothing
    ElseIf InStr(UCase(TypeName(obj)), "ISITEDOUBLE") <> 0 Then
        SD_temp_obj = obj
        If gDictSiteDouble.Exists(keyname) Then
            gDictSiteDouble.Remove (keyname)
        End If
        gDictSiteDouble.Add keyname, SD_temp_obj
        Set SD_temp_obj = Nothing
    ElseIf InStr(UCase(TypeName(obj)), "ISITELONG") <> 0 Then
        SL_temp_obj = obj
        If gDictSiteLong.Exists(keyname) Then
            gDictSiteLong.Remove (keyname)
        End If
        gDictSiteLong.Add keyname, SL_temp_obj
        Set SL_temp_obj = Nothing
    ElseIf InStr(UCase(TypeName(obj)), "IDSPWAVE") <> 0 Then
        DSPWF_temp_obj = obj
        If gDictDSPWaves.Exists(keyname) Then
            gDictDSPWaves.Remove (keyname)
        End If
        gDictDSPWaves.Add keyname, DSPWF_temp_obj
        '20220106, Add for Real value validation( reverse bit)
        If gB_efuse_DicValue_Chk_Flag = True Then ''Efuse_DicValue_Chk
            For Each Vsite In theexec.sites
                If DSPWF_temp_obj.SampleSize = 1 Then
                    objvalue = DSPWF_temp_obj.ConvertDataTypeTo(DspLong)
                    theexec.Datalog.WriteComment "Site(" + CStr(Vsite) + ")" + "@@@ Key_Name:= " + keyname + " Value:= " + CStr(objvalue.Element(0))
                Else
                    objvalue = DSPWF_temp_obj.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, DSPWF_temp_obj.SampleSize, 0)
                    theexec.Datalog.WriteComment "Site(" + CStr(Vsite) + ")" + "@@@ Key_Name:= " + keyname + " Value:= " + CStr(objvalue.Element(0))
                End If
            Next Vsite
        End If
        ''Efuse_DicValue_Chk <==
        Set DSPWF_temp_obj = Nothing
    Else
        Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "StoreDataAllType", "[" & keyname & "]Input data type is not identify!!")
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "AddStoredData") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function GetStoreDataAllType(keyname As String) As Variant
    'ReStore data from Dictionary for Multi-Data Type -- 20230627
    'DataType : SiteDouble/SiteLong/PinListData/DSPWave
    keyname = LCase(keyname)
    If gDictSiteLong.Exists(keyname) Then
        Set GetStoreDataAllType = gDictSiteLong(keyname)
    ElseIf gDictSiteDouble.Exists(keyname) Then
        Set GetStoreDataAllType = gDictSiteDouble(keyname)
    ElseIf gDictDSPWaves.Exists(keyname) Then
        Set GetStoreDataAllType = gDictDSPWaves(keyname)
    ElseIf gDictCurrMeasurements.Exists(keyname) Then
        Set GetStoreDataAllType = gDictCurrMeasurements(keyname)
    Else
        Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "GetStoreDataAllType", "[" & keyname & "] Not found!!")
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Dictionary", "GetStoreDataAllType") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
