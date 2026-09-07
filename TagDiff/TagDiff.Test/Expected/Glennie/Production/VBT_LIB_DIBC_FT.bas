Attribute VB_Name = "VBT_LIB_DIBC_FT"
Option Explicit

' const
Public Const DIBC_FOLDER = "dibc_thera_ft"
'Public F_TDR_Recal As Boolean
Private Const VERSION = "V2021.09.28"
Private Const DOUBLE_MAX = 1E+304
Private Const DOUBLE_MIN = -1E+304
Private Const COVERAGE_INFO_OPEN_ONLY = ""
Private Const DIBC_DEBUG = False  ' Should be False while releasing

' DIBC.ini
Private disableDIBC        As Boolean
Private ignorePassFlag     As Boolean
Private failStop           As Boolean
Private testShortItems     As Boolean
Private logFolder          As String

' DIBC_Limits.csv
Private loLimitDict        As Dictionary
Private hiLimitDict        As Dictionary
Private tNameDict          As Dictionary

' DIBC_Testability.csv
Private tNumOpenColl       As Collection
Private tNumShortColl      As Collection
Private isShortDict        As Dictionary

' DIBC_Channel.csv
Private pinSiteChannelDict As Dictionary

' result row
Private Type DIBCResultRow
    tNum        As String
    site        As String
    Tname       As String
    pin         As String
    channel     As String
    component   As String
    low         As Double
    measured    As Double
    high        As Double
    unit        As String
    isPass      As Boolean
End Type

' result row str
Private Type DIBCResultRowStr
    tNum        As String
    site        As String
    Tname       As String
    pin         As String
    channel     As String
    component   As String
    lowStr      As String
    measuredStr As String
    highStr     As String
End Type

' pin meas data
Private Type PinMeasData
   PinName      As String
   Threshold    As New SiteDouble
   MeasVal      As New SiteDouble
   MeasPF       As New SiteLong
   FinishSearch As New SiteBoolean
End Type

' expect data
Private Type ExecData
   CurrMeasData()  As PinMeasData
   PrevUpperData() As PinMeasData
   PrevLowerData() As PinMeasData
End Type

' global variable
Private StartTime            As Date
Private dibcPassFlag         As Boolean
Private allLogRows           As Collection
Private failLogRows          As Collection
Private unitScaleDict        As Dictionary
Private dibcResultPerSite    As Dictionary
Private dibcPassCountPerType As Dictionary
Private dibcAllCountPerType  As Dictionary
Private dibonTesterType As String '' add for UltraFLEXplus or Jaguar 211228

' relayOn pins for digital open test
Private Relay_On_Pin_Dic_open        As New Dictionary
Private Relay_On_Pin_Dic_short       As New Dictionary
Private Relay_On_Pin_Dic_short_UVI80 As New Dictionary


Public g_ReadPinMap As Boolean
Public g_map_PinGroup As Object
' ----------------------------------------------------------------
' Name:    DIBC_Start
' Purpose: DIBC main entrance
' ----------------------------------------------------------------
Public Sub FT_DIBC_Start()

    On Error GoTo DIBC_Error

    Call DIBC_Init
    dibonTesterType = TheHdw.tester.type ' add for UltraFLEXplus or Jaguar 211228
    ' entry filter
    If Not Util_Is_TDR_Passed() Then  ' TDR pass
        Debug.Print "TDR failed, bypass DIBC"

        For Each site In TheExec.sites
            TheExec.sites.item(site).SortNumber = 7
            TheExec.sites.item(site).BinNumber = 7
            TheExec.sites.item(site).result = tlResultFail
        Next site


        Exit Sub

    ElseIf disableDIBC Then  ' not disable DIBC
        Debug.Print "Leave DIBC without execution due to DisableDIBC Flag set to True"
        Exit Sub

    ElseIf (Not ignorePassFlag And dibcPassFlag) Then  ' not DIBC passed before
        Debug.Print "Leave DIBC without execution due to already DIBC already passed and IgnorePassFlag set to False"
        Exit Sub

    End If

    ' dibc body
    Call DIBC_Test
    Call Util_Write_Log
    Call DIBC_DisconnectAll
    Call Util_Write_Log
    Call DIBC_Discharge
    Call DIBC_DisconnectAll
    
    TheHdw.Alarms.Check
    
    For Each site In TheExec.sites
        TheExec.sites.item(site).SortNumber = 7
        TheExec.sites.item(site).BinNumber = 7
        TheExec.sites.item(site).result = tlResultFail
    Next site

    ' fail stop
    If (Not dibcPassFlag) And failStop Then Call DIBC_Fail_Retry_Or_Abort


    On Error GoTo 0
    Exit Sub

DIBC_Error:
    MsgBox "Error " & err.number & " (" & err.Description & ") in procedure DIBC_Start of Sub DIBC"
End Sub
Private Sub DIBC_DisconnectAll()

   Dim strAllDigitalPins() As String
    Dim strAllDCVSPins() As String

    Dim strAllDCVSPins2() As String
    Dim strAllDCVSPins3() As String
    Dim strAllDCVSPins4() As String
    Dim strAllDCVSPins5() As String

    Dim strAllDCVIPins() As String
    
    Dim strAllUtilsPins() As String
    Dim lngPnum As Long
    
    Dim sAllDigitalPins As String
    Dim sAllDCVSPins As String
    
    Dim sAllDCVSPins2 As String
    Dim sAllDCVSPins3 As String
    Dim sAllDCVSPins4 As String
    Dim sAllDCVSPins5 As String
    
    Dim sAllDCVIPins As String
    
    Dim sAllUtilsPins As String
    
    ' All Digital Pin
    Call TheExec.DataManager.GetPinNames(strAllDigitalPins, chIO, lngPnum)
    sAllDigitalPins = Join(strAllDigitalPins, ",")
    
    DIBC_PPMU_DisConnect (sAllDigitalPins)
'    thehdw.Digital.DisConnectPins (sAllDigitalPins)
    TheHdw.Digital.pins(sAllDigitalPins).Disconnect
    
    ' All DCVS Pin
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins, chDCVS, lngPnum)
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins2, chDCVSMerged2, lngPnum)
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins3, chDCVSMerged4, lngPnum)
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins4, chDCVSMerged6, lngPnum)
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins5, chDCVSMerged8, lngPnum)
    
    
    sAllDCVSPins = Join(strAllDCVSPins, ",")
    sAllDCVSPins2 = Join(strAllDCVSPins2, ",")
    sAllDCVSPins3 = Join(strAllDCVSPins3, ",")
    sAllDCVSPins4 = Join(strAllDCVSPins4, ",")
    sAllDCVSPins5 = Join(strAllDCVSPins5, ",")
    
    sAllDCVSPins = sAllDCVSPins & "," & sAllDCVSPins2 & "," & sAllDCVSPins3 & "," & sAllDCVSPins5
    
    DIBC_DCVS_DisConnect (sAllDCVSPins)
    
    ' All DCVI Pin
    ''' Remove for Donan @William 230518
''    Call TheExec.DataManager.GetPinNames(strAllDCVIPins, chDCVI, lngPnum)
''    sAllDCVIPins = Join(strAllDCVIPins, ",")
''
''    DIBC_DCVI_DisConnect (sAllDCVIPins)
    
    ' All Utility Pin
    Call TheExec.DataManager.GetPinNames(strAllUtilsPins, chUtil, lngPnum)
    sAllUtilsPins = Join(strAllUtilsPins, ",")
    TheHdw.Utility.pins(sAllUtilsPins).State = tlUtilBitOff
    
End Sub

Private Sub DIBC_Fail_Retry_Or_Abort()
    On Error GoTo errHandler

    Dim ans As Integer
    ans = MsgBox("DIBC detected some fail items." + vbNewLine + "Press 'OK' to execute DIBC again", _
                 vbOKOnly + vbExclamation, _
                 Title:="DIBC Fail")

        '230414 Updated
    If ans = vbOK Then
        Do While TheHdw.DIB.IsConnected = False
            ans = MsgBox("Detected DIB undock.", vbOKOnly + vbExclamation, "DIB undock")
        Loop

        If TheHdw.DIB.IsConnected = True Then
            F_TDR_Recal = True
        End If
    End If
    Exit Sub

errHandler:
    HandleExecIPError "Error Happened On Retry TDR Calibrated"

End Sub

' ----------------------------------------------------------------
' Procedure Name: DIBC_Test
' Purpose: Perform tests according to the testability csv
' ----------------------------------------------------------------
Private Sub DIBC_Test()

    On Error GoTo DIBC_Test_Error
    

    'Call ResetAllInstrument
    'Call SmartRelaySwitch(DGS_Relay)
    TheHdw.Wait 1
    ' set DIB power-on
    TheHdw.DIB.powerOn = True
    TheHdw.Wait 0.001

    ' test items in openNum
    Dim tNum As Variant
    If DIBC_DEBUG Then
        Debug.Print "[WARNING] DIBC in DEBUG Mode..."
        For Each tNum In tNumOpenColl
            DIBC_Test_Debug CStr(tNum)
        Next tNum
    Else
        For Each tNum In tNumOpenColl
            DIBC_Test_Try_Invoke CStr(tNum)
            If TheExec.sites.ActiveCount = 0 Then GoTo Fail_handler
        Next tNum
    End If

    ' wait for contact wafer
    If testShortItems And tNumShortColl.Count > 0 Then

        Dim ans As Integer
        ans = MsgBox("Press 'OK' after metal wafer is contacted" + vbNewLine + _
                     "Press 'Cancel' to skip short items", _
                     vbOKCancel + vbQuestion, _
                     Title:="Contact Metal Wafer")

        If ans <> vbCancel Then
            ' test items in shortNum
            For Each tNum In tNumShortColl
                DIBC_Test_Try_Invoke CStr(tNum)
                If TheExec.sites.ActiveCount = 0 Then GoTo Fail_handler
            Next tNum
        End If

    End If

    ' set DIB power-off
    TheHdw.DIB.powerOn = False
    TheHdw.Wait 0.001

    ' update global pass flag
    dibcPassFlag = (Util_Site_Summary(-1, True) = "PASS")
    Debug.Print "DIBC " + Util_Site_Summary(-1, True)

    On Error GoTo 0
    Exit Sub
    
Fail_handler:
    TheHdw.DIB.powerOn = False
    TheHdw.Wait 0.001
    Exit Sub
    

DIBC_Test_Error:
    TheHdw.DIB.powerOn = False
    MsgBox "Error " & err.number & " (" & err.Description & ") in procedure DIBC_Test of Sub DIBC"
End Sub

Private Sub DIBC_Test_Debug(tNum As String)

    Debug.Print "DIBC Test: " + tNum

    Dim argarray As Variant: argarray = Array("A", "B", "C")

    Dim meas As PinListData: Set meas = Util_Debug_Pseudo_Data(argarray)

    Util_Judge_Result meas, tNum, "PATH", ""

End Sub

Private Sub DIBC_Test_Try_Invoke(tNum As String)

    On Error GoTo DIBC_Test_Try_Invoke_Error

    Run "FT_DIBC_Test_" + tNum

    On Error GoTo 0
    Exit Sub

DIBC_Test_Try_Invoke_Error:
    Debug.Print "Error " & err.number & " (" & err.Description & ") in procedure DIBC_Test_Try_Invoke " & tNum & " of Sub DIBC"
End Sub

' ----------------------------------------------------------------
' Name:    Read_<FileNames>
' Purpose: Set global variables from files
' ----------------------------------------------------------------
Private Sub Read_DIBC_Ini()

    Dim fileNum As Long:    fileNum = freefile
    Dim FilePath As String: FilePath = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\DIBC.ini"
    Open FilePath For Input As fileNum

        Dim line As String
        Dim tokens() As String

        Do Until EOF(fileNum)

            Line Input #fileNum, line
            tokens = Split(line, "=")

            If UBound(tokens) = 1 Then
                Dim key As String: key = Trim(UCase(tokens(0)))
                Dim val As String: val = tokens(1)

                Select Case key
                    Case "DISABLEDIBC"
                        disableDIBC = Util_Is_True(val)

                    Case "IGNOREPASSFLAG"
                        ignorePassFlag = Util_Is_True(val)

                    Case "FAILSTOP"
                        failStop = Util_Is_True(val)

                    Case "TESTSHORTITEMS"
                        testShortItems = Util_Is_True(val)

                    Case "LOGFOLDER"
                        logFolder = Trim(val)

                End Select
            End If
        Loop
    Close fileNum

End Sub

Private Sub Read_DIBC_Limit()

    ' reset variable
    Set loLimitDict = New Dictionary
    Set hiLimitDict = New Dictionary
    Set tNameDict = New Dictionary

    ' read file
    Dim fileNum As Long:    fileNum = freefile
    Dim FilePath As String: FilePath = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\DIBC_Limits.csv"

    Open FilePath For Input As fileNum
        Dim line          As String
        Dim tokens()      As String
        Dim lineNum       As Long
        Dim tNameOverflow As Boolean

        Do Until EOF(fileNum)

            Line Input #fileNum, line
            lineNum = lineNum + 1

            If lineNum <> 1 Then
                tokens = Split(line, ",")
                If UBound(tokens) >= 5 Then

                    Dim tNum    As String: tNum = Trim(tokens(0))
                    Dim Tname   As String: Tname = Trim(tokens(1))
                    Dim pin     As String: pin = Trim(tokens(2))
                    Dim lowStr  As String: lowStr = Trim(tokens(3))
                    Dim highStr As String: highStr = Trim(tokens(4))
                    'Dim unit    As String: unit = Trim(tokens(5))

                    If IsNumeric(lowStr) Then loLimitDict(tNum & "_" & Trim(UCase(pin))) = CDbl(lowStr)
                    If IsNumeric(highStr) Then hiLimitDict(tNum & "_" & Trim(UCase(pin))) = CDbl(highStr)

                    tNameDict(tNum) = Tname
                    If Len(Tname) >= 32 And Not tNameOverflow Then
                        tNameOverflow = True
                        MsgBox "tName Overflow in DIBC_Limit.csv: " + Tname
                    End If

                End If
            End If
        Loop
    Close fileNum

End Sub

Private Sub Read_DIBC_Testability()

    ' reset variable
    Set tNumOpenColl = New Collection
    Set tNumShortColl = New Collection
    Set isShortDict = New Dictionary

    ' read file
    Dim fileNum  As Long:    fileNum = freefile
    Dim FilePath As String:  FilePath = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\DIBC_Testability.csv"

    Open FilePath For Input As fileNum

        Dim line     As String
        Dim tokens() As String
        Dim lineNum  As Long

        Do Until EOF(fileNum)

            Line Input #fileNum, line
            lineNum = lineNum + 1

            If lineNum <> 1 Then ' ignore header
                tokens = Split(line, ",")
                If UBound(tokens) >= 3 And Not Util_Is_True(tokens(0)) Then  ' not disable

                    Dim isShort As Boolean: isShort = (UCase(Trim(tokens(1))) = "SHORT")
                    Dim tNum    As String:  tNum = Trim(tokens(3))

                    If isShort Then
                        tNumShortColl.Add tNum
                    Else
                        tNumOpenColl.Add tNum
                    End If

                End If
            End If
        Loop
    Close fileNum

End Sub

Private Sub Read_DIBC_Channel()

    ' reset variable
    Set pinSiteChannelDict = New Dictionary

    ' read file
    Dim fileNum  As Long:   fileNum = freefile
    Dim FilePath As String: FilePath = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\DIBC_Channel.csv"

    Open FilePath For Input As fileNum

        Dim line As String
        Dim tokens() As String
        Dim lineNum As Long

        Do Until EOF(fileNum)
            Line Input #fileNum, line
            lineNum = lineNum + 1

            If lineNum <> 1 Then  ' ignore header
                tokens = Split(line, ",")
                If UBound(tokens) >= 2 Then
                    Dim PinName As String:    PinName = UCase(Trim(tokens(0)))
                    Dim channelStr As String: channelStr = tokens(2)
                    Dim channelPogoStr As String: channelPogoStr = tokens(3)
                    pinSiteChannelDict(Trim(UCase(PinName))) = Split(channelPogoStr, "+")
                End If
            End If
        Loop

    Close fileNum

End Sub


' ----------------------------------------------------------------
' Name:    DIBC_Init
' Purpose: Init global variables and read setting files
' ----------------------------------------------------------------
Private Sub DIBC_Init()

    StartTime = Now()

    Set allLogRows = New Collection
    Set failLogRows = New Collection

    Set dibcResultPerSite = New Dictionary
    Dim site As Variant
    For Each site In TheExec.sites.Active
        dibcResultPerSite(CLng(site)) = True  ' default pass
    Next site
    dibcResultPerSite(-1) = True  ' all sites

    Set dibcAllCountPerType = New Dictionary
    Set dibcPassCountPerType = New Dictionary

    Set unitScaleDict = New Dictionary

    unitScaleDict.Add -15, "f"
    unitScaleDict.Add -12, "p"
    unitScaleDict.Add -9, "n"
    unitScaleDict.Add -6, "u"
    unitScaleDict.Add -3, "m"
    unitScaleDict.Add 0, ""
    unitScaleDict.Add 3, "K"
    unitScaleDict.Add 6, "M"
    unitScaleDict.Add 9, "G"
    unitScaleDict.Add 12, "T"

    ' read files
    Call Read_DIBC_Ini
    Call Read_DIBC_Limit
    Call Read_DIBC_Testability
    Call Read_DIBC_Channel

End Sub

Private Function Util_Is_True(val As String) As Boolean

    Util_Is_True = (UCase(Trim(val)) = "TRUE")

End Function

Private Sub Util_Judge_Result(meas As PinListData, tNum As String, component As String, unit As String)

    Dim site As Variant
    Dim pin As Variant

    For Each site In TheExec.sites.Active
        For Each pin In meas.pins

            Dim loLim    As Double
            Dim hiLim    As Double
            Util_Get_Limits tNum, CStr(pin), loLim, hiLim

            Dim measured As Double: measured = meas.pins(pin).value(site)
            Dim Row      As DIBCResultRow

            Row.tNum = tNum
            Row.site = CStr(site)
            Row.Tname = tNameDict(tNum)
            Row.pin = CStr(pin)
            Row.channel = Util_Get_Channel(CStr(pin), CLng(site))
            Row.component = component
            Row.low = loLim
            Row.measured = measured
            Row.high = hiLim
            Row.unit = unit

            'row.isPass = loLim <= measured And measured <= hiLim
            If loLim <> DOUBLE_MIN And hiLim <> DOUBLE_MAX Then
                Row.isPass = loLim <= measured And measured <= hiLim
            ElseIf loLim = DOUBLE_MIN Then
                Row.isPass = measured <= hiLim
            ElseIf hiLim = DOUBLE_MAX Then
                Row.isPass = loLim <= measured
            End If

            Util_Log_Result Row
                        'Ben Modify
            Call TheExec.flow.TestLimit(resultVal:=measured, ForceResults:=tlForceNone, unit:=unitCustom, customUnit:=unit, Tname:=Row.Tname, _
            lowVal:=loLim, hiVal:=hiLim, PinName:=Row.pin)
        Next
    Next

End Sub

Private Sub Util_Get_Limits(tNum As String, pin As String, loLim As Double, hiLim As Double)

    Dim key1 As String: key1 = tNum & "_" & Trim(UCase(pin))
    Dim key2 As String: key2 = tNum & "_"

    If loLimitDict.Exists(key1) Then
        loLim = loLimitDict(key1)
    ElseIf loLimitDict.Exists(key2) Then
        loLim = loLimitDict(key2)
    Else
        loLim = DOUBLE_MIN
    End If

    If hiLimitDict.Exists(key1) Then
        hiLim = hiLimitDict(key1)
    ElseIf hiLimitDict.Exists(key2) Then
        hiLim = hiLimitDict(key2)
    Else
        hiLim = DOUBLE_MAX
    End If

End Sub

Private Sub Util_Log_Result(Row As DIBCResultRow)

    Dim rowStr   As DIBCResultRowStr:  rowStr = Util_Result_Row_To_Str(Row)
    Dim PrintStr As String:            PrintStr = Util_Result_Row_To_Print_Str(rowStr)
    allLogRows.Add PrintStr

    If Not Row.isPass Then
        failLogRows.Add PrintStr
        dibcResultPerSite(CLng(Row.site)) = False
        dibcResultPerSite(-1) = False
    End If

    ' counter
    If Not dibcAllCountPerType.Exists(Row.component) Then
        dibcPassCountPerType(Row.component) = 0
        dibcAllCountPerType(Row.component) = 0
    End If

    dibcAllCountPerType(Row.component) = dibcAllCountPerType(Row.component) + 1
    If Row.isPass Then
        dibcPassCountPerType(Row.component) = dibcPassCountPerType(Row.component) + 1
    End If

End Sub

Private Function Util_Debug_Pseudo_Data(pins As Variant) As PinListData

    Dim pld  As New PinListData
    Dim pin  As Variant

    For Each pin In pins
        pld.AddPin (CStr(pin))
    Next pin

    Set Util_Debug_Pseudo_Data = pld

End Function

Private Function Util_Get_Channel(PinName As String, site As Long) As String


    On Error GoTo Util_Get_Channel_Error

    Util_Get_Channel = "-1"  ' default value
    Util_Get_Channel = pinSiteChannelDict(Trim(UCase(PinName)))(site)

    On Error GoTo 0
    Exit Function

Util_Get_Channel_Error:
    ' do nothing
End Function

Private Function Util_Result_Row_To_Print_Str(Row As DIBCResultRowStr) As String

    Dim Fields(0 To 8) As String
    Fields(0) = Util_Left_Padding(Row.tNum, 10)
    Fields(1) = Util_Left_Padding(Row.site, 7)
    Fields(2) = Util_Left_Padding(Row.Tname, 33)
    Fields(3) = Util_Left_Padding(Row.pin, 33)
        'New format for DGS Channel. from 12 -> 13
    Fields(4) = Util_Left_Padding(Row.channel, 13)
    Fields(5) = Util_Left_Padding(Row.component, 24)
    Fields(6) = Util_Left_Padding(Row.lowStr, 17)
    Fields(7) = Util_Left_Padding(Row.measuredStr, 16)
    Fields(8) = Util_Left_Padding(Row.highStr, 16)
    Util_Result_Row_To_Print_Str = Join(Fields, "")

End Function

Private Function Util_Left_Padding(val As String, length As Integer, Optional Character As String = " ")

    Dim padStr As String
    Dim i      As Integer
    For i = 1 To length
        padStr = Character & padStr
    Next i
    padStr = padStr & Trim(val)
    Util_Left_Padding = right(padStr, length)

End Function

Private Sub Util_Write_Log()

    If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder  ' todo: check if parent folder exists

    Dim fileNum As Long:    fileNum = freefile
    Dim logPath As String:  logPath = logFolder & "/" & Util_Log_Filename

    Open logPath For Output As fileNum

        ' header
        Print #fileNum, "DIB Checker Datalog Report: " & VERSION
        Print #fileNum, Format(StartTime, "mm/dd/yyyy hh:mm:ss AM/PM")
        Print #fileNum, Util_Left_Padding("Program Name:", 16) & "    " & TheExec.TestProgram.name
        Print #fileNum, Util_Left_Padding("Tester Name:", 16) & "    " & Environ("COMPUTERNAME")
        Print #fileNum, Util_Left_Padding("DIB Serial:", 16) & "    " & Util_Get_Dib_Serial()
        Print #fileNum, vbNewLine & Util_Left_Padding("", 50, "=") & vbNewLine
        Print #fileNum, Util_Header()

        ' rows
        Dim Row As Variant
        If failLogRows.Count > 0 Then
            For Each Row In failLogRows
                Print #fileNum, Row
            Next Row
            Print #fileNum, ""
        End If

        For Each Row In allLogRows
            Print #fileNum, Row
        Next Row

        ' summary
        Print #fileNum, vbNewLine & "Summary of DIB Checker Results"
        Print #fileNum, Util_Left_Padding("", 50, "=")
        Print #fileNum, Util_Left_Padding("Site", 10) & Util_Left_Padding("P/F", 6)

        Dim site As Variant
        For Each site In TheExec.sites.Active
            Print #fileNum, Util_Left_Padding(CStr(site), 10) & Util_Left_Padding(Util_Site_Summary(CLng(site)), 6)
        Next site

        Print #fileNum, "Overall DIB Checker Result : " & Util_Site_Summary(-1, True)

        Print #fileNum, vbNewLine & "DIB checker results :"
        Print #fileNum, Util_Left_Padding("", 50, "=")
        Print #fileNum, Util_Pass_Rate()

        ' footer
        Print #fileNum, vbNewLine & "Total Execution time : " & CStr(DateDiff("s", StartTime, Now())) & ".0s"

        Print #fileNum, vbNewLine & "Test coverage information"
        Print #fileNum, Util_Left_Padding("", 50, "=")
        Print #fileNum, COVERAGE_INFO_OPEN_ONLY
    Close fileNum

End Sub

Private Function Util_Pass_Rate() As String

    Dim passRateStr As String: passRateStr = Util_Left_Padding("", 5) & "Pass Rate : "
    Dim passRate As Double

    If allLogRows.Count = 0 Then
        passRate = 1#
    Else
        passRate = (allLogRows.Count - failLogRows.Count) / allLogRows.Count
    End If

    passRateStr = passRateStr & Format(passRate, "0.00%")
    passRateStr = passRateStr & " (Pass Number: " & CStr(allLogRows.Count - failLogRows.Count) & "; Total Number: " & CStr(allLogRows.Count) & ")"

    Dim component As Variant
    For Each component In dibcAllCountPerType
        passRate = dibcPassCountPerType(component) / dibcAllCountPerType(component)
        passRateStr = passRateStr & vbNewLine & Util_Left_Padding("", 10) & CStr(component) & " : " & Format(passRate, "0.00%")
    Next component

    Util_Pass_Rate = passRateStr

End Function

Private Function Util_Site_Summary(site As Long, Optional verbose As Boolean = False) As String

    If Not dibcResultPerSite.Exists(site) Then
        Util_Site_Summary = "U"
        Exit Function
    End If

    Util_Site_Summary = IIf(dibcResultPerSite(site), IIf(verbose, "PASS", "P"), IIf(verbose, "FAIL", "F"))

End Function

Private Function Util_Header() As String

    Dim Row As DIBCResultRowStr
    Row.tNum = "Number"
    Row.site = "Site"
    Row.Tname = "Test Name"
    Row.pin = "Pin"
    Row.channel = "Channel"
    Row.component = "Fail Components"
    Row.lowStr = "Low"
    Row.measuredStr = "Measured"
    Row.highStr = "High"
    Util_Header = Util_Result_Row_To_Print_Str(Row)

End Function

Private Function Util_Result_Row_To_Str(Row As DIBCResultRow) As DIBCResultRowStr

    Dim rowStr As DIBCResultRowStr
    rowStr.tNum = Row.tNum
    rowStr.site = Row.site
    rowStr.pin = Row.pin
    rowStr.Tname = Row.Tname
    rowStr.channel = Row.channel
    rowStr.component = Row.component
    rowStr.lowStr = IIf(Row.low = DOUBLE_MIN, "", Util_Align_Unit(Row.low, Row.unit))
    rowStr.measuredStr = Util_Align_Unit(Row.measured, Row.unit, Row.isPass)
    rowStr.highStr = IIf(Row.high = DOUBLE_MAX, "", Util_Align_Unit(Row.high, Row.unit))
    Util_Result_Row_To_Str = rowStr

End Function

Private Function Util_Is_TDR_Passed() As Boolean

    Util_Is_TDR_Passed = True

    If DIBC_DEBUG Then Exit Function

    Dim site As Variant
    For Each site In TheExec.sites.Existing
        If dibonTesterType = "Jaguar" Then
            If TheHdw.Digital.Calibration.site(site).Status = tlCalSiteStatusFail Then
                Util_Is_TDR_Passed = False
                Exit For
            End If
        ElseIf dibonTesterType = "UltraFLEXplus" Then
            If TheHdw.Calibration.Status = tlCalibrationStatus_Fail Then
                Util_Is_TDR_Passed = False
                Exit For
            End If
        End If
    Next site

End Function

Private Function Util_Align_Unit(val As Double, unit As String, Optional isPass As Boolean = True) As String

    Dim result As String
    Dim valWithScale As Double:  valWithScale = val
    Dim power        As Integer: power = 0

    While Abs(valWithScale) >= 99
        valWithScale = valWithScale / 1000#
        power = power + 3
    Wend

    While Abs(valWithScale) <= 0.099 And Abs(valWithScale) > 1E-18 ' avoid valWithScale = 0 causing infinite loop
        valWithScale = valWithScale / 0.001
        power = power - 3
    Wend

    Dim prefix As String
    If unitScaleDict.Exists(power) Then
        prefix = unitScaleDict(power)
        result = Format(valWithScale, "0.000") + " " + prefix + unit
    Else
        result = Format(valWithScale, "Scientific") + " " + prefix + unit
    End If

    If Not isPass Then result = "(F)" + Util_Left_Padding(result, 12)

    Util_Align_Unit = result

End Function

Private Function Util_Get_Dib_Serial() As String

    On Error GoTo Util_Get_Dib_Serial_Error

    Util_Get_Dib_Serial = "DUMMY"

    'Const PROBE_CARD_ID_REG_ADDR = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\PROBECARD_ID"

    With CreateObject("wscript.shell")
        Util_Get_Dib_Serial = .RegRead("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\LOADBOARD_ID")
    End With

    On Error GoTo 0
    Exit Function

Util_Get_Dib_Serial_Error:
    Debug.Print "Error " & err.number & " (" & err.Description & ") in procedure Util_Get_Dib_Serial of Function DIBC"
End Function

Private Function Util_Log_Filename() As String

    Dim tokens(5) As String
    tokens(0) = "DIBCheckerResult"
    tokens(1) = Split(TheExec.TestProgram.name, ".")(0)              ' tpName without extension
    tokens(2) = Util_Get_Dib_Serial()                                ' DIB Serial
    tokens(3) = Environ("COMPUTERNAME")                              ' tester
    tokens(4) = Format(StartTime, "yyyymmdd-hhmmAM/PM")              ' timestamp
    tokens(5) = IIf(failLogRows.Count > 0, "Fail.txt", "Pass.txt")   ' pass / fail flag
    Util_Log_Filename = Join(tokens, "_")

End Function

Private Function Util_Incr(tNum As String, incr As Long) As String
    Util_Incr = CStr(CLng(tNum) + incr)
End Function

' ********************************************************************************
' =========                  DIBC Unit Test Suites                       =========
' ********************************************************************************

Private Sub Test_Suites()
    DIBC_Init
    Test_Util_Align_Unit
    Test_Pin_Site_Channel_Dict
    Test_Util_Incr
End Sub

Private Sub Test_Util_Align_Unit()
    Debug.Assert Util_Align_Unit(0.00012345, "A") = "0.123 mA"
    Debug.Assert Util_Align_Unit(-0.00012345, "A") = "-0.123 mA"
    Debug.Assert Util_Align_Unit(0.000012345, "A") = "12.345 uA"
    Debug.Assert Util_Align_Unit(12345, "Ohm", False) = "(F) 12.345 KOhm"
    Debug.Print "Test_Util_Align_Unit: PASS"
End Sub

Private Sub Test_Pin_Site_Channel_Dict()
    Debug.Assert Util_Get_Channel("XI0", 0) <> "-1"
    Debug.Print "Test_Pin_Site_Channel_Dict: PASS"
End Sub

Private Sub Test_Util_Incr()
    Debug.Assert Util_Incr("100", 10) = "110"
    Debug.Print "Test_Util_Incr: PASS"
End Sub

Private Sub DIBC_Discharge()

    Dim strAllDigitalPins() As String
    Dim strAllDCVSPins() As String
    
    Dim strAllDCVSPins2() As String
    Dim strAllDCVSPins3() As String
    Dim strAllDCVSPins4() As String
    Dim strAllDCVSPins5() As String
    
    Dim strAllDCVIPins() As String
    Dim strAllDCVIPins2() As String
    
    Dim strAllUtilsPins() As String
    Dim lngPnum As Long
    
    Dim sAllDigitalPins As String
    
    Dim sAllDCVSPins As String
    Dim sAllDCVSPins2 As String
    Dim sAllDCVSPins3 As String
    Dim sAllDCVSPins4 As String
    Dim sAllDCVSPins5 As String
    
    Dim sAllDCVIPins As String
    Dim sAllDCVIPins2 As String
    
    Dim sAllUtilsPins As String
     
    ' All Digital Pin
    Call TheExec.DataManager.GetPinNames(strAllDigitalPins, chIO, lngPnum)
    'sAllDigitalPins = Join(strAllDigitalPins, ",")
    
    sAllDigitalPins = GetAllPinByChannlType("I/O")
    
    
    DIBC_PPMU_Set0V (sAllDigitalPins)
'    TheHdw.Digital.DisConnectPins (sAllDigitalPins)
    
    ' All DCVS Pin
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins, chDCVS, lngPnum)
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins2, chDCVSMerged2, lngPnum)
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins3, chDCVSMerged4, lngPnum)
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins4, chDCVSMerged6, lngPnum)
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins5, chDCVSMerged8, lngPnum)
    
    sAllDCVSPins = Join(strAllDCVSPins, ",")
    sAllDCVSPins2 = Join(strAllDCVSPins2, ",")
    sAllDCVSPins3 = Join(strAllDCVSPins3, ",")
    sAllDCVSPins4 = Join(strAllDCVSPins4, ",")
    sAllDCVSPins5 = Join(strAllDCVSPins5, ",")
    
    
    'sAllDCVSPins = sAllDCVSPins & "," & sAllDCVSPins2 & "," & sAllDCVSPins3 & "," & sAllDCVSPins5
    
    sAllDCVSPins = GetAllPinByChannlType("DCVS")
    
    sAllDCVSPins = Replace(sAllDCVSPins, ",PMU_IREF", "")
    
    DIBC_DCVS_Set0V (sAllDCVSPins)
    
    ' All DCVI Pin
    'Call TheExec.DataManager.GetPinNames(strAllDCVIPins, chDCVI, lngPnum)
    'Call TheExec.DataManager.GetPinNames(strAllDCVIPins2, chDCVIMerged, lngPnum)
    
'    sAllDCVIPins = Join(strAllDCVIPins, ",")
'    sAllDCVIPins2 = Join(strAllDCVIPins2, ",")
'
'    sAllDCVIPins = sAllDCVIPins & "," & sAllDCVIPins2
'
'    DIBC_DCVI_Set0V (sAllDCVIPins)

    
End Sub

Private Function GetAllPinByChannlType(chaType As String) As String
On Error GoTo errHandler
    Dim currChannelMap As String: currChannelMap = LCase(TheExec.CurrentChanMap)
    Dim chanmap_content() As Variant
    Dim i As Long
    Dim j As Long
    Dim indx As Long
    Dim Count As Long: Count = 0
    Dim max_row As Long
    Dim max_col As Long
    Dim Pin_row As Long
    Dim Pin_col As Long
    Dim Type_col As Long
    Dim pin_name As String: pin_name = "Pin Name"
    Dim Type_Name As String: Type_Name = "Type"
    Dim pinStr As String
    Dim chaTypePattern As Object
    Set chaTypePattern = CreateObject("vbscript.regexp")
    
    Dim pinStrDict As New Dictionary

    chaTypePattern.Pattern = "^" & chaType
    chaTypePattern.IgnoreCase = False
    
    Worksheets(currChannelMap).Activate
    max_row = Worksheets(currChannelMap).UsedRange.Rows.Count
    max_col = Worksheets(currChannelMap).UsedRange.Columns.Count
    chanmap_content = Worksheets(currChannelMap).range(Cells(1, 1), Cells(max_row, max_col)).value
    
    For i = 1 To max_row
        For j = 1 To max_col - 1
            If chanmap_content(i, j) <> "" Then
                If chanmap_content(i, j) = pin_name Then '''Get the position of "Pin Name" from channel map
                    Pin_row = i
                    Pin_col = j
                ElseIf chanmap_content(i, j) = Type_Name Then '''Get the position of "Type" from channel map
                    Type_col = j
                    Exit For
                End If
            End If
        Next j
    Next i

    For i = Pin_row + 1 To max_row
    
        'If chanmap_content(i, Type_col) = UCase(chaType) Then
        If chaTypePattern.test(chanmap_content(i, Type_col)) Then
            If pinStr = "" Then
                pinStr = chanmap_content(i, Pin_col)
            End If
            pinStr = Replace(chanmap_content(i, Pin_col), ":C", "")
            pinStr = Replace(chanmap_content(i, Pin_col), ":M", "")
            
            If Not pinStrDict.Exists(Replace(chanmap_content(i, Pin_col), ":C", "")) Then
                pinStrDict.Add Replace(chanmap_content(i, Pin_col), ":C", ""), chaType
            End If
            'pinStr = pinStr & "," & Replace(chanmap_content(i, Pin_col), ":C", "")
        End If
    
    Next i
    
    pinStr = Join(pinStrDict.Keys, ",")
    
    GetAllPinByChannlType = pinStr
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_GetAllPinByChannlType"
End Function


'Central Ben Edit
'230502 Disconnect instrument
' ----------------------------------------------------------------
' Name:    DIBC_PPMU_DisConnect
' Purpose: Disconnect All PPMU instrument
' ----------------------------------------------------------------
Public Sub DIBC_PPMU_DisConnect(strPins As String)
                Dim reg As Object
    ' delete unused ","
    Set reg = CreateObject("vbscript.regexp")
    reg.Pattern = "^(\s*\,)+"
    strPins = reg.Replace(strPins, "")
    reg.Pattern = "(\,\s*)+$"
    strPins = reg.Replace(strPins, "")
    reg.Pattern = "(\,\s*\,)+"
    strPins = reg.Replace(strPins, ",")
    If Not strPins = "" Then
        With TheHdw.PPMU.pins(strPins)
            .Gate = tlOff
             TheHdw.Wait 0.002
            .Reset tlResetSettings
            .Disconnect
        End With
    End If
End Sub
Public Sub DIBC_PPMU_Set0V(strPins As String)
                Dim reg As Object
    ' delete unused ","
    Set reg = CreateObject("vbscript.regexp")
    reg.Pattern = "^(\s*\,)+"
    strPins = reg.Replace(strPins, "")
    reg.Pattern = "(\,\s*)+$"
    strPins = reg.Replace(strPins, "")
    reg.Pattern = "(\,\s*\,)+"
    strPins = reg.Replace(strPins, ",")
    If Not strPins = "" Then
        With TheHdw.PPMU.pins(strPins)
            .Gate = tlOff
             TheHdw.Wait 0.002
            .ForceV 0
            .Connect
            .Gate = tlOn
             
 '          .Reset tlResetSettings

        End With
    
    TheHdw.Wait 5
    End If
End Sub

' ----------------------------------------------------------------
' Name:    DIBC_DCVI_DisConnect
' Purpose: Disconnect All DCVI instrument
' ----------------------------------------------------------------
Public Sub DIBC_DCVI_DisConnect(strPins As String, Optional Instrument As String)
    If strPins <> "" Then
        With TheHdw.DCVI.pins(strPins)
            .Alarm(tlDCVIAlarmAll) = tlAlarmDefault
            .Gate = False
            TheHdw.Wait 0.002
            .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 0.0005
        
        If UCase(Instrument) = "DC-07" Then
        ' When the DCVI is disconnected from the DIB, the bleeder resistor default connection state is connected (on).
        'it helps quickly discharge DIB capacitance.
            TheHdw.DCVI.pins(strPins).BleederResistor = tlDCVIBleederResistorOn   'only UVI80 have bleedresistor
            TheHdw.Wait 0.001
        End If
    End If
End Sub
Public Sub DIBC_DCVI_Set0V(strPins As String, Optional Instrument As String)
    If strPins <> "" Then
        With TheHdw.DCVI.pins(strPins)
            .Alarm(tlDCVIAlarmAll) = tlAlarmDefault
            .Gate = False
            .mode = tlDCVIModeVoltage
            .Voltage = 0
            .Connect tlDCVIConnectDefault
            
             TheHdw.Wait 0.002
             .Gate = True
'            .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 5
        
        If UCase(Instrument) = "DC-07" Then
        ' When the DCVI is disconnected from the DIB, the bleeder resistor default connection state is connected (on).
        'it helps quickly discharge DIB capacitance.
            TheHdw.DCVI.pins(strPins).BleederResistor = tlDCVIBleederResistorOn   'only UVI80 have bleedresistor
            TheHdw.Wait 0.001
        End If
    End If
End Sub
' ----------------------------------------------------------------
' Name:    DIBC_DCVS_DisConnect
' Purpose: Disconnect All DCVS instrument
' ----------------------------------------------------------------
Public Sub DIBC_DCVS_DisConnect(strPins As String)
    If Not strPins = "" Then
        With TheHdw.DCVS.pins(strPins)
            .CurrentRange.value = 0.02
            .Voltage.Main = 0#
            .Voltage.Alt = 0#
            TheHdw.Wait 0.005
            .Alarm(tlDCVSAlarmAll) = tlAlarmDefault
            .Gate = False
            TheHdw.Wait 0.002
            .Disconnect tlDCVSConnectDefault
        End With
        TheHdw.Wait 0.0005
    End If
End Sub

Public Sub DIBC_DCVS_Set0V(strPins As String)
    
 '   TheHdw.Utility.pins("K86").State = tlUtilBitOn
    If Not strPins = "" Then
        With TheHdw.DCVS.pins(strPins)
            .CurrentRange.value = 0.02
            .Voltage.Main = 0
            .Voltage.Alt = 0
            TheHdw.Wait 0.005
            .Alarm(tlDCVSAlarmAll) = tlAlarmDefault
            .Connect
            TheHdw.Wait 0.002
            .Gate = True
            
           End With
        TheHdw.Wait 5
    End If
End Sub
'Private Sub DIBC_Fail_Retry_Or_Abort()
'    On Error GoTo ErrHandler
'
'    Dim ans As Integer
'    ans = MsgBox("DIBC detected some fail items." + vbNewLine + "Press 'OK' to execute DIBC again", _
'                 vbOKOnly + vbExclamation, _
'                 Title:="DIBC Fail")
'
'        '230414 Updated
'    If ans = vbOK Then
'        Do While TheHdw.DIB.IsConnected = False
'            ans = MsgBox("Detected DIB undock.", vbOKOnly + vbExclamation, "DIB undock")
'        Loop
'
'        If TheHdw.DIB.IsConnected = True Then
'            F_TDR_Recal = True
'        End If
'    End If
'    Exit Sub
'
'ErrHandler:
'    HandleExecIPError "Error Happened On Retry TDR Calibrated"
'
'End Sub

' ********************************************************************************
' =========               Auto Generated DIBC_Lib                        =========
' ********************************************************************************

Private Function FT_DIBC_Lib_MeasureIOleakage(tNum As String, argarray As Variant) As PinListData
    
    On Error GoTo errHandler

    Dim strAllDigitalPins() As String
    Dim strAllDCVIPins() As String
    Dim strAllDCVSPins() As String
    Dim strAllDCVSMerged2Pins() As String
    Dim strAllDCVSMerged3Pins() As String
    Dim strAllDCVSMerged4Pins() As String
    Dim strAllDCVSMerged5Pins() As String
    Dim strAllDCVSMerged6Pins() As String
    Dim strAllDCVSMerged7Pins() As String
    Dim strAllDCVSMerged8Pins() As String
    Dim strAllDCVSMerged10Pins() As String
    Dim strAllDCVSMerged16Pins() As String
    Dim sAllDigitalPins As String
    Dim sAllDCVIPins As String
    Dim sAllDCVSPins As String
    Dim sAllDCVSMerged2Pins As String
    Dim sAllDCVSMerged3Pins As String
    Dim sAllDCVSMerged4Pins As String
    Dim sAllDCVSMerged5Pins As String
    Dim sAllDCVSMerged6Pins As String
    Dim sAllDCVSMerged7Pins As String
    Dim sAllDCVSMerged8Pins As String
    Dim sAllDCVSMerged10Pins As String
    Dim sAllDCVSMerged16Pins As String
    
    Dim lngPnum As Long
    Dim pldMeasureValuePPMUOFF As New PinListData
    Dim pldMeasureValuePPMUON As New PinListData
    Dim pldMeasureValuePPMU As New PinListData
    Dim pldIOMeasure As New PinListData
    Dim sPin As Variant
    Dim thissite As Variant
    Dim str_IO_loopback As String
    Dim lIndex As Long
    Dim i As Integer
    'Dim Exclude_Arr() As Variant
    
    'Exclude_Arr = Split(argArray, ",")
    
    
    'Call SmartRelaySwitch(DGS_Relay)
    
        
    
    Call TheExec.DataManager.GetPinNames(strAllDigitalPins, chIO, lngPnum)
        
    'sAllDigitalPins = sAllDigitalPins = GetAllPinByChannlType("I/O")
    sAllDigitalPins = Join(strAllDigitalPins, ",")
'    For i = 0 To UBound(argArray)
'    sAllDigitalPins = Replace(UCase(sAllDigitalPins), argArray(i) & ",", "")
'    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "," & argArray(i), "")
'    Next i
    
    
    
    If sAllDigitalPins = "" Then
        Exit Function
    End If
    
    strAllDigitalPins = Split(sAllDigitalPins, ",")
    
    If sAllDigitalPins = "" Then
        Exit Function
    Else
        With TheHdw.PPMU.pins(sAllDigitalPins)
                .Connect
                TheHdw.Wait 0.0005
                .ForceV 0, 0.002
                .Gate = tlOn
                TheHdw.Wait 0.005
        End With

    End If
        'sAllDCVSPins = GetAllPinByChannlType("DCVS")
    Call TheExec.DataManager.GetPinNames(strAllDCVSPins, chDCVS, lngPnum)
    sAllDCVSPins = Join(strAllDCVSPins, ",")
    Call TheExec.DataManager.GetPinNames(strAllDCVSMerged2Pins, chDCVSMerged2, lngPnum)
    sAllDCVSMerged2Pins = Join(strAllDCVSMerged2Pins, ",")
    'Call TheExec.DataManager.GetPinNames(strAllDCVSMerged3Pins, chDCVSMerged3, lngPnum)
    'sAllDCVSMerged3Pins = Join(strAllDCVSMerged3Pins, ",")
    Call TheExec.DataManager.GetPinNames(strAllDCVSMerged4Pins, chDCVSMerged4, lngPnum)
    sAllDCVSMerged4Pins = Join(strAllDCVSMerged4Pins, ",")
    'Call TheExec.DataManager.GetPinNames(strAllDCVSMerged5Pins, chDCVSMerged5, lngPnum)
    'sAllDCVSMerged5Pins = Join(strAllDCVSMerged5Pins, ",")
    Call TheExec.DataManager.GetPinNames(strAllDCVSMerged6Pins, chDCVSMerged6, lngPnum)
    sAllDCVSMerged6Pins = Join(strAllDCVSMerged6Pins, ",")
    'Call TheExec.DataManager.GetPinNames(strAllDCVSMerged7Pins, chDCVSMerged7, lngPnum)
    'sAllDCVSMerged7Pins = Join(strAllDCVSMerged7Pins, ",")
     Call TheExec.DataManager.GetPinNames(strAllDCVSMerged8Pins, chDCVSMerged8, lngPnum)
    sAllDCVSMerged8Pins = Join(strAllDCVSMerged8Pins, ",")
    'Call TheExec.DataManager.GetPinNames(strAllDCVSMerged10Pins, chDCVSMerged10, lngPnum)
    'sAllDCVSMerged10Pins = Join(strAllDCVSMerged10Pins, ",")
    'Call TheExec.DataManager.GetPinNames(strAllDCVSMerged16Pins, chDCVSMerged16, lngPnum)
    'sAllDCVSMerged16Pins = Join(strAllDCVSMerged16Pins, ",")
    sAllDCVSPins = sAllDCVSPins & "," & sAllDCVSMerged2Pins & "," & sAllDCVSMerged4Pins & "," & sAllDCVSMerged6Pins & "," & sAllDCVSMerged8Pins
  
  
      If Not sAllDCVSPins = "" Then
        'set all DCVS pins to 0
        TheHdw.DCVS.pins(sAllDCVSPins).Connect tlDCVSConnectDefault
        TheHdw.Wait 0.0005
        With TheHdw.DCVS.pins(sAllDCVSPins)
          .mode = tlDCVSModeVoltage
          .Voltage.Main.value = 0
          .Voltage.Alt.value = 0
          .Voltage.Output = tlDCVSVoltageMain
          .CurrentRange.value = 0.2
          .Meter.mode = tlDCVSMeterVoltage
          '.VoltageRange = 2
          .Meter.Filter.bypass = False
          .Alarm(tlDCVSAlarmAll) = tlAlarmOff
          .Gate = True
          TheHdw.Wait 0.01
        End With
    End If
  

'    If Not sAllDCVSPins = "" Then
'        thehdw.DCVS.pins(sAllDCVSPins).Disconnect
'        thehdw.Wait 0.0005
'
'        With thehdw.DCVS.pins(sAllDCVSPins)
'            .mode = tlDCVSModeHighImpedance
'            .Gate = False
'            thehdw.Wait 0.01
'        End With
'    End If
'

    For Each sPin In strAllDigitalPins
       If ExistsInPinGroup(sPin, "TDR_Exclude_Pins") = 1 Then
       
        Else
        
            '***set PPMU instrument***'
'            With TheHdw.PPMU.pins(sPin)
'                    .Disconnect
'                    TheHdw.Wait 0.001
'                    .ClampVHi = 5 * v
'                    .Gate = tlOn
'                    TheHdw.Wait 0.003
'                    .ForceV 5, 0.002
'                    TheHdw.Wait 0.002
'                    .ForceV 5, 0.000002
'                    TheHdw.Wait 0.005
'            End With
   
                        
'            Set pldMeasureValuePPMUOFF = TheHdw.PPMU.pins(sPin).Read(tlPPMUReadMeasurements, 128, tlPPMUReadingFormatAverage)
                        
                        '230414 Update
            '***connect PPMU instrument***'
            With TheHdw.PPMU.pins(sPin)
                .Connect
                    TheHdw.Wait 0.002
                                        '.ClampVHi (6.5)
                    .ForceV 1, 0.002
                    TheHdw.Wait 0.002
            End With
            TheHdw.Wait 0.005
            
            Set pldMeasureValuePPMUON = TheHdw.PPMU.pins(sPin).Read(tlPPMUReadMeasurements, 128, tlPPMUReadingFormatAverage)
'            Set pldMeasureValuePPMU = pldMeasureValuePPMUON.Math.Abs.Subtract(pldMeasureValuePPMUOFF.Math.Abs)
            Set pldMeasureValuePPMU = pldMeasureValuePPMUON
            
           pldIOMeasure.AddPin sPin
           pldIOMeasure.pins(sPin) = pldMeasureValuePPMU.pins(sPin)
           '*** PPMU force 0v ***'
            With TheHdw.PPMU.pins(sPin)
                    .ForceV 0, 0.000002
                    TheHdw.Wait 0.005
            End With

        End If
        
    Next sPin

    '***disconnect PPMU instrument***'
    If Not sAllDigitalPins = "" Then
        With TheHdw.PPMU.pins(sAllDigitalPins)
        
                .Gate = tlOff
                TheHdw.Wait 0.002
                .Reset tlResetSettings
                .Disconnect
        End With
     End If

    '***disconnect DCVS instrument***'
    If Not sAllDCVSPins = "" Then
        With TheHdw.DCVS.pins(sAllDCVSPins)
          .CurrentRange.value = 0.02
          .Voltage.Main = 0#
          .Voltage.Alt = 0#
          .Gate = False
           TheHdw.Wait 0.002
          .Disconnect tlDCVSConnectDefault
        End With
        TheHdw.Wait 0.0005
    End If
  ' Set DIBC_Lib_MeasureIOleakage = pldIOMeasure
   
        Util_Judge_Result pldIOMeasure, 4000, "IO_Leakage", "A"
        
        If TheExec.sites.ActiveCount = 0 Then Exit Function

        '********************main test end********************'
        'Call SmartRelaySwitch("")
   
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_MeasureIOleakage"
End Function

'Ben Modify 20220303
Private Sub FT_DIBC_Lib_UtilityCheck(tNum As String, argarray As Variant)


'Ben Modify 20220307
'Cory Modify 20220307
        '********************udb check main test begin********************'
        On Error GoTo errHandler

        Dim arrAllRelayPins() As String
        Dim strRelayPins_tested As String
        Dim pldMeasureValue As New PinListData
        Dim i As Integer
        Dim iStartNum As Integer
        Dim lngPnum As Long
        
        Call TheExec.DataManager.GetPinNames(arrAllRelayPins, chUtil, lngPnum)
        
        strRelayPins_tested = ""
        For i = 0 To UBound(arrAllRelayPins)
                strRelayPins_tested = strRelayPins_tested & arrAllRelayPins(i) & ", "
                
                '20(max) relays as one group,since support borad max current
                If (i + 1) Mod 20 = 0 Or i = UBound(arrAllRelayPins) Then
                        strRelayPins_tested = left(strRelayPins_tested, Len(strRelayPins_tested) - 1) 'remove last ", "
                        
                        If (i + 1) Mod 20 = 0 Then
                                iStartNum = i - 19
                        ElseIf i = UBound(arrAllRelayPins) Then
                                iStartNum = i - i Mod 20
                        End If
                        
                        'relay on check
                        Set pldMeasureValue = DIBC_Lib_CheckUdbState(strRelayPins_tested, "ON")
                        
                        Call TheExec.flow.TestLimit(resultVal:=pldMeasureValue, ForceResults:=tlForceNone, unit:=unitCustom, customUnit:="", Tname:="UDBCHECK_ON", _
                        lowVal:=0, hiVal:=0)
                        Util_Judge_Result pldMeasureValue, 4100, "UDBCHECK_ON", ""
                        
                        If TheExec.sites.ActiveCount = 0 Then Exit Sub
                        
                        'relay off check
                        Set pldMeasureValue = DIBC_Lib_CheckUdbState(strRelayPins_tested, "OFF")
                        
                        Call TheExec.flow.TestLimit(resultVal:=pldMeasureValue, ForceResults:=tlForceNone, unit:=unitCustom, customUnit:="", Tname:="UDBCHECK_OFF", _
                        lowVal:=1, hiVal:=1)
                        Util_Judge_Result pldMeasureValue, 4100 + 10, "UDBCHECK_ON", ""
                        If TheExec.sites.ActiveCount = 0 Then Exit Sub
                        
                        Set pldMeasureValue = Nothing
                        strRelayPins_tested = ""
                End If
        Next i
        '********************main test end********************'
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UtilityCheck"
End Sub


Private Sub FT_DIBC_Lib_SPI_ROM(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Call Read_Status_Reg(tNum, argarray)

    Call Read_Device_ID(CStr(CLng(tNum) + 10), argarray)

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_SPI_ROM"
End Sub













Private Sub FT_DIBC_Lib_UP1600_IO(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Call DIBC_Lib_UP1600_IO_Open(tNum, argarray)

    Call Digital_Open_Voltage_Meas(CStr(CLng(tNum) + 10), argarray, "")
   
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP1600_IO"
End Sub


Private Sub FT_DIBC_Lib_UP1600_R(tNum As String, argarray As Variant)
End Sub
Private Sub FT_DIBC_Lib_UP1600_RELAY(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray = Array("RT_CLK32768","RT_CLK32768_PA","K01")

     Dim RLY_name As String: RLY_name = argarray(2)
    Dim Test_Pin As String: Test_Pin = argarray(0)

    Call DIBC_measureRelayTDRDelta_2(tNum, RLY_name, Test_Pin)
    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), RLY_name, Test_Pin)

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP1600_RELAY"
End Sub


Private Sub FT_DIBC_Lib_UP1600_RELAY_R(tNum As String, argarray As Variant)
On Error GoTo errHandler
Dim Relay_On As String
Dim Route_Start As String
Dim Route_End As String
Dim Rly_temp As String
Dim TestPin_temp As String
Dim TestPin_ary() As String
Dim i As Integer
Route_Start = argarray(0)
Relay_On = argarray(1)
Call Relay_Circuit_check(tNum, 0.05, 0.004, 1.7, 2.3, Relay_On, 0.003, Route_Start, "")
   
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP1600_RELAY_R"
End Sub




Private Sub FT_DIBC_Lib_UP1600_RELAY_R_2(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray = Array("ATC1_USB_RESREF","K32","K33")
    Dim p1 As String: p1 = argarray(0)
    Dim k1 As String: k1 = argarray(1)
    Dim k2 As String: k2 = argarray(2)

    ' toggle relay (k1: OFF, k2: OFF)
    TheHdw.Utility.pins(k1).State = tlUtilBitOff
    TheHdw.Utility.pins(k2).State = tlUtilBitOff

    Call DIBC_measureRelayTDRDelta_2(tNum, k1, p1)

    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), k1, p1)

    ' toggle relay (K1: OFF, k2: ON)
    TheHdw.Utility.pins(k1).State = tlUtilBitOff
    TheHdw.Utility.pins(k2).State = tlUtilBitOn

    Call Relay_Circuit_check(CStr(CLng(tNum) + 20), 0.05, 0.004, 1.7, 2.3, k1, 0.003, p1, "")

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP1600_RELAY_R_2"
End Sub

Private Sub FT_DIBC_Lib_UP1600_RELAY_2(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray = Array("XI0", "XI0_PA", "XO0", "XO0_PA", "K2")
    Dim p1 As String: p1 = argarray(0)
    Dim p2 As String: p2 = argarray(2)
    Dim k1 As String: k1 = argarray(4)

    Call DIBC_measureRelayTDRDelta_2(tNum, k1, p1)
    Call DIBC_measureRelayTDRDelta_2(tNum, k1, p2)

    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), k1, p1)
    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), k1, p2)
   
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP1600_RELAY_2"
End Sub

Private Sub FT_DIBC_Lib_UVI80_RELAY(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray = Array("ANALOGMUX_OUT", "K1")
    Dim p1 As String: p1 = argarray(0)
    Dim k1 As String: k1 = argarray(1)
    
    Call UVI80_Leakage(tNum, p1, k1)
    
    Call DCVI_RLY_Rising_Meas(CStr(CLng(tNum) + 10), p1, "", k1)
  
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVI80_RELAY"
End Sub


Private Sub FT_DIBC_Lib_UVI80_RELAY_C(tNum As String, argarray As Variant)
On Error GoTo errHandler
'argArray = Array("PAD_MTR_ANALOG_TEST_N","K10","1.20e-06")
    Dim Relay_On As String
    Dim Test_PinName As String
    
    Relay_On = argarray(1)
    Test_PinName = argarray(0)
    Call DCVI_Cap_meas_1(tNum, Test_PinName, Relay_On)
  
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVI80_RELAY_C"
End Sub
Private Sub FT_DIBC_Lib_UVI80_Relay_C_2(tNum As String, argarray As Variant)
On Error GoTo errHandler
    'argArray = Array("PAD_MTR_ANALOG_TEST_N","K10", "K11", "1.20e-06")

    Dim pin     As String: pin = argarray(0)
    Dim relayOn As String: relayOn = argarray(1) + "," + argarray(2)
    
    Call DCVI_Cap_meas_1(tNum, pin, relayOn)
  
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVI80_Relay_C_2"
End Sub

Private Sub FT_DIBC_Lib_UVI80_RELAY_C_3(tNum As String, argarray As Variant)
On Error GoTo errHandler
    'argArray = Array("PAD_MTR_VREF_P","PAD_MTR_VREF_N","K13","K12","1.00e-06")

    Dim p1 As String: p1 = argarray(0)
    Dim p2 As String: p2 = argarray(1)
    Dim k1 As String: k1 = argarray(2)
    Dim k2 As String: k2 = argarray(3)
    Dim relayOn As String: relayOn = Join(Array(k1, k2), ",")

    Call DCVI_Cap_Meas_3(tNum, p1, p2, relayOn)

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVI80_RELAY_C_3"
End Sub


Private Sub FT_DIBC_Lib_UVI80_RELAY_R(tNum As String, argarray As Variant)
On Error GoTo errHandler
    'argArray = Array("ANALOGMUX_OUT_SRC","K09","K07","K08")
    Dim p1 As String: p1 = argarray(0)
    Dim k1 As String: k1 = argarray(1)
    Dim k2 As String: k2 = argarray(2)
    Dim k3 As String: k3 = argarray(3)

    Call Relay_Circuit_check(tNum, 6, 0.004, 1.7, 2.3, Join(Array(k2, k3), ","), 0.003, p1, "")

    Call DCVI_RLY_Rising_Meas(Util_Incr(tNum, 10), p1, "", k1)

    TheHdw.Utility.pins(k1).State = tlUtilBitOn
    Call DCVI_RLY_Rising_Meas(Util_Incr(tNum, 20), p1, "", k2)

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVI80_RELAY_R"
End Sub

Private Sub FT_DIBC_Lib_UVS256_RELAY_C_3(tNum As String, argarray As Variant)
On Error GoTo errHandler
    'argArray = Array("PAD_MTR_VREF_P","PAD_MTR_VREF_N","K13","K12","1.00e-06")

    Dim p1 As String: p1 = argarray(0)
    Dim p2 As String: p2 = argarray(1)
    Dim k1 As String: k1 = argarray(2)
    Dim k2 As String: k2 = argarray(3)
    Dim relayOn As String: relayOn = Join(Array(k1, k2), ",")

    Call DCVI_Cap_Meas_3(tNum, p1, p2, relayOn)

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVS256_RELAY_C_3"
End Sub


Private Sub FT_DIBC_Lib_UVS256_C(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray =  Array("VDD12_ADC:4.15e-05","VDD12_AMUX:3.92e-05", ...)
    Dim tokens() As String
    Dim i        As Integer
    For i = 0 To UBound(argarray)
        tokens = Split(argarray(i), ":")  ' assert UBound(tokens) = 1
        Call DIBC_Lib_HEXVS_C_Open_MeasureCapacitor_2(tNum, tokens(0), "VHDVS", Format(tokens(1), "general number"), 0.15, 1.2, -1, -1, 0)
    Next i

    tNum = CStr(CLng(tNum) + 10)
    For i = 0 To UBound(argarray)
        tokens = Split(argarray(i), ":")
        Call UVS256_HexVs_Leakage(tNum, 0.1, 0, 0.000004, tokens(0), "UVS256")
    Next i
     
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVS256_C"
End Sub
Private Sub FT_DIBC_Lib_UVS256_RELAY_UP1600(tNum As String, argarray As Variant)
On Error GoTo errHandler
    'argArray = Array("VDDQL_DDR", "DDR0_RREF", "K31", "DDR1_RREF", "K32", "DDR2_RREF", "K33", "DDR3_RREF", "K34")
    Dim powerPin As String: powerPin = argarray(0)
    Dim p1       As String
    Dim k1       As String

    Dim i As Integer
    For i = 1 To UBound(argarray) Step 2
        If i + 1 >= UBound(argarray) Then Exit For
        p1 = argarray(i)
        k1 = argarray(i + 1)
        Call Relay_Circuit_check(tNum, 0.05, 0.004, 1.7, 2.3, k1, 0.003, powerPin, p1)
    Next i

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVS256_RELAY_UP1600"
End Sub
Private Sub FT_DIBC_Lib_HEXVS_C(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray = Array("VDD_CPU_SRAM:6.95e-04","VDD_DCS_DDR:7.59e-04", ...)
    Dim tokens() As String
    Dim i        As Integer

    ' cap val
    For i = 0 To UBound(argarray)
        tokens = Split(argarray(i), ":")  ' assert ubound(tokens) == 1
        Call DIBC_Lib_HEXVS_C_Open_MeasureCapacitor_2(tNum, tokens(0), "HEXVS", Format(tokens(1), "general number"), 0.15, 1.2, -1, -1, 0)
    Next i

    ' cap leak
    tNum = CStr(CLng(tNum) + 10)
    For i = 0 To UBound(argarray)
        tokens = Split(argarray(i), ":")
        Call UVS256_HexVs_Leakage(tNum, 0.1, 0, 0.000004, tokens(0), "HEXVS")
    Next i

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_HEXVS_C"
End Sub

Private Sub FT_DIBC_Lib_COMPLEX(tNum As String, argarray As Variant)
On Error GoTo errHandler
'argArray = Array("VDDQL_DDR", "DDR0_RREF", "K31", "DDR1_RREF", "K32", "DDR2_RREF", "K33", "DDR3_RREF", "K34")
    Dim k1 As String
    Dim k2 As String
    Dim K9 As String
    Dim K10 As String
    Dim K11 As String
    Dim K12 As String
    Dim Route_Start As String
    Dim Route_End As String
    Dim Relay_On As String
    Dim UVI80_pin As String
    Dim Rly_temp As String
    Dim TestPin_temp As String
    Dim TestPin_ary() As String
    Dim RlyPin_ary() As String
    Dim SubCKT_Num As Integer
    Dim FT_SubCKT_Num As Integer
    Dim FT_SubCKT_idx As Integer
    Dim Component_idx As Integer:: Component_idx = 3
    ReDim TestPin_ary(3) As String
    ReDim RlyPin_ary(3) As String
    Dim i As Integer
    SubCKT_Num = argarray(2)
    Route_Start = argarray(0)
    For i = 1 To SubCKT_Num
        UVI80_pin = argarray(Component_idx)
        Component_idx = Component_idx + 1
        Route_End = argarray(Component_idx)
        Component_idx = Component_idx + 1
        K9 = argarray(Component_idx)
        Component_idx = Component_idx + 1
        K10 = argarray(Component_idx)
        Component_idx = Component_idx + 1
        K11 = argarray(Component_idx)
        Component_idx = Component_idx + 1
        K12 = argarray(Component_idx)
        Component_idx = Component_idx + 1
        Relay_On = K9 & "," & K10
        Call Relay_Circuit_check(tNum, 0.05, 0.004, 1.7, 2.3, Relay_On, 0.003, Route_Start, Route_End)
        tNum = CStr(CLng(tNum) + 10)
        Call DIBC_measureRelayTDRDelta_2(tNum, K10, Route_End)
        tNum = CStr(CLng(tNum) + 10)
        Call DIBC_PPMU_Relay_FIMV_rising_new(tNum, K10, Route_End)
        tNum = CStr(CLng(tNum) + 10)
        Call DCVI_RLY_Rising_Meas(tNum, UVI80_pin, "", K12)
        tNum = CStr(CLng(tNum) + 10)
        TheHdw.Utility.pins(K12).State = tlUtilBitOn
        Call DCVI_RLY_Rising_Meas(tNum, UVI80_pin, "", K11)
        TheHdw.Utility.pins(K12).State = tlUtilBitOff
    Next i
        FT_SubCKT_idx = Component_idx
        FT_SubCKT_Num = argarray(FT_SubCKT_idx)
        Component_idx = Component_idx + 1
    If LCase(TheExec.CurrentJob) Like "*ft*" Then
        For i = 1 To FT_SubCKT_Num
            Route_End = argarray(Component_idx)
            Component_idx = Component_idx + 1
            k1 = argarray(Component_idx)
            Component_idx = Component_idx + 1
            k2 = argarray(Component_idx)
            Component_idx = Component_idx + 1
            Call Relay_Circuit_check(tNum, 0.05, 0.004, 1.7, 2.3, Relay_On, 0.003, Route_Start, Route_End)
            tNum = CStr(CLng(tNum) + 10)
            Call DIBC_measureRelayTDRDelta_2(tNum, k2, Route_End)
            tNum = CStr(CLng(tNum) + 10)
            Call DIBC_PPMU_Relay_FIMV_rising_new(tNum, k2, Route_End)
        Next i
    End If

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_COMPLEX"
End Sub
Private Sub FT_DIBC_Lib_DIFF_RELAY(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray := Array(pin1, pin2, K1, K2, K3, K4, C)
    Dim p1 As String: p1 = argarray(0)
    Dim p2 As String: p2 = argarray(1)
    Dim k1 As String: k1 = argarray(2)
    Dim k2 As String: k2 = argarray(3)
    Dim k3 As String: k3 = argarray(4)
    Dim k4 As String: k4 = argarray(5)

    Call DIBC_measureRelayTDRDelta_2(tNum, k1, p1)
    
    Call DIBC_PPMU_Relay_FIMV_rising_new(Util_Incr(tNum, 10), k1, p1)
    
    Call DIBC_measureRelayTDRDelta_2(Util_Incr(tNum, 20), k4, p2)
    
    Call DIBC_PPMU_Relay_FIMV_rising_new(Util_Incr(tNum, 30), k4, p2)
    
    Call Relay_Circuit_check(Util_Incr(tNum, 40), 0.05, 0.004, 1.7, 2.3, Join(Array(k1, k2, k3, k4), ","), 0.003, p1, p2)
    
    Call Relay_Parasitic_Caps(Util_Incr(tNum, 50), p2, k1, k2, k3, k4)

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_DIFF_RELAY"
End Sub


'Private Sub FT_DIBC_Lib_FRC_BUFFER(tNum As String, argarray As Variant)
'On Error GoTo ErrHandler
'    Dim PortName  As String: PortName = "XI0_Diff_Port"
'    Dim DUTPin    As String: DUTPin = "XI0_PA"
'    Dim Bufferpin As String: Bufferpin = "REFCLK_RT_CLK32768"
'
'    ' clock Fan-out Buffer
'    Call FreeRunClk_Disable_dibchecker("XI0_Diff_Port") 'disable clock
'
'    With TheHdw.Digital.pins(PortName).Levels
'        .value(chVil) = 0  'm_sAllDigitalPinList
'        .value(chVih) = 1.8
'        .value(chIol) = 0
'        .value(chIoh) = 0
'        .value(chVch) = 6
'        .value(chVoutLoTyp) = 0
'        .value(chVoutHiTyp) = 0
'        .DriverMode = tlDriverModeLargeHiZ  'Largeswing-HiZ
'    End With
'
'    With TheHdw.Digital.pins(DUTPin).Levels
'        .value(chVol) = 0.54  'dut pin
'        .value(chVoh) = 0.54
'        .value(chVt) = 0.9
'        .value(chVcl) = 0
'    End With
'
'
'    TheHdw.Digital.pins(Bufferpin).Levels.value(chVol) = 0.55
'    TheHdw.Digital.pins(Bufferpin).Levels.value(chVoh) = 0.55
'    TheHdw.Digital.pins(Bufferpin).Levels.value(chVt) = 0.6
'    TheHdw.Digital.pins(Bufferpin).Levels.value(chVcl) = -1
'    TheHdw.Wait 0.01
'    TheHdw.Digital.pins(Bufferpin).Connect
'    TheHdw.Wait 0.005
'
'    ' set support board
'    Dim SBFreq As Double:  SBFreq = 24 * MHz
'    Dim SB_Vih As Double:  SB_Vih = 1.8
'    With TheHdw.DIB.SupportBoardClock
'        .Connect
'        .Frequency = SBFreq
'        .Vih = SB_Vih ' Max is 6V
'        .Vil = 0      ' Min is -1V
'        .start
'    End With
'
'    Call MeasFreq_dibchecker(tNum, Bufferpin, 0.001, "REFCLK")
'
'    Call FreeRunclk_Enable_dibchecker(tNum + 10, PortName, DUTPin, Bufferpin, "REFCLK_XIO_PLLLock")
'
'    ' clean up
'    TheHdw.Digital.pins(Bufferpin).Disconnect
'    Call FreeRunClk_Disable_dibchecker("XI0_Diff_Port") 'disable clock
'
'Exit Sub
'ErrHandler:
'    TheExec.Datalog.WriteComment "error in DIBC_Lib_FRC_BUFFER"
'End Sub






Private Sub FT_DIBC_Lib_LOOP_BACK(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray = Array("ST_PCIE_RX1_N","ST_PCIE_TX1_N","1.00e-07")
    Dim p1 As String: p1 = argarray(0)
    Dim p2 As String: p2 = argarray(1)

    Call DIBC_Lib_LOOP_BACK_IO_Leakage_Open(tNum, Array(p1, p2))

    Call DIBC_Lib_LoopBack_1MHz_VOH_check(CStr(CLng(tNum) + 10), Array(p1, p2))

    Call DIBC_Lib_LoopBack_Freq(CStr(CLng(tNum) + 20), Array(p1, p2))

    Call DIBC_Lib_Loop_Back_MCC_Cap_meas(CStr(CLng(tNum) + 30), Array(p1, p2), 0.02)
     
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_LOOP_BACK"
End Sub

Private Sub FT_DIBC_Lib_SB_EEPROM(tNum As String, argarray As Variant)
On Error GoTo errHandler
        Dim REC() As IDIB_EEPROM_BlockObj
        ' EEPROMREC = iEEPROM.Block.List
         '  EEPROMREC = thehdw.DIB.EEPROM.Block.List
        REC = TheHdw.DIB.EEPROM.Block.list

        Dim tmpBoardNumber As Double
        tmpBoardNumber = REC(0).ID
        'tmpBoardNumber = thehdw.DIB.EEPROM.Block.List.ID
        Dim tmpBoardSize As Double
        tmpBoardSize = REC(0).size
       ' tmpBoardSize = thehdw.DIB.EEPROM.Block.List.size
        Dim vsite As Variant
        Dim pldMeasureValue As New PinListData
        Dim ResultsValue As Double
        Dim Tname As String
        Tname = "EEPROMRead"
           
        If tmpBoardNumber >= 0 And tmpBoardSize >= 0 Then
            ResultsValue = 1
        Else
            ResultsValue = 0
        End If
        
        pldMeasureValue.AddPin " "  'edit by rita on 2018May23'
        For Each vsite In TheExec.sites.Active '0521
            pldMeasureValue.pins(" ").value(vsite) = ResultsValue 'edit by rita on 2018May23'
        Next vsite '0521
       ' judgeCheckResult pldMeasureValue, 8004, "EEPROMRead", 1, 1
 'TheExec.Flow.TestLimit resultVal:=pldMeasureValue, lowVal:=1, hiVal:=1, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
        Util_Judge_Result pldMeasureValue, tNum, "DIBROM_Check", ""


Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_SB_EEPROM"
End Sub
Private Sub FT_DIBC_Lib_SB_Power(tNum As String, argarray As Variant)
On Error GoTo errHandler
    'header
    Debug.Print "Test " + tNum
    Dim Power_List() As String
    ' reset DIB power
    TheHdw.DIB.powerOn = False
    TheHdw.Wait 0.5
    Power_List = TheHdw.DIB.power.list
    ' loop through each power
    Dim PowerValue  As Variant
    Dim powerStr    As String
    
    For Each PowerValue In TheHdw.DIB.power.list
        powerStr = UCase(CStr(PowerValue))
        
        Select Case powerStr
            Case "30.UPS48V_A", "31.UPS48V_A", "48V"
                Call Util_Meas_SB_Power(tNum, PowerValue, 48)
                
            Case "30.UPS15V_A", "30.UPS15V_B", "31.UPS15V_A", "31.UPS15V_B", "15V_2", "15V_1"
                Call Util_Meas_SB_Power(tNum, PowerValue, 15)
                
            Case "30.UPS12V_A", "31.UPS12V_A", "12V"
                Call Util_Meas_SB_Power(tNum, PowerValue, 12)
                
            Case "30.UPS5V_A", "30.UPS5V_B", "31.UPS5V_A", "31.UPS5V_B", "5V_2", "5V_1"
                Call Util_Meas_SB_Power(tNum, PowerValue, 5)
            
            Case "30.UPS3_P3V_A", "31.UPS3_P3V_A"
                Call Util_Meas_SB_Power(tNum, PowerValue, 3)
            
            Case "3.3V"
                Call Util_Meas_SB_Power(tNum, PowerValue, 3.3)
                
            Case Else
                Debug.Print "SB power not handled: " + powerStr
                
        End Select
        'tNum = CStr(CLng(tNum) + 1)
    Next PowerValue

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_SB_Power"
End Sub


Private Sub FT_DIBC_Lib_UP2200_IO(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Call DIBC_Lib_UP2200_IO_Open(tNum, argarray)

    Call Digital_Open_Voltage_Meas(CStr(CLng(tNum) + 10), argarray, "")
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP2200_IO"
End Sub

Private Sub FT_DIBC_Lib_TmpSensor(tNum As String, argarray As Variant)
End Sub
'Ben Modify 230605
Private Sub FT_DIBC_Lib_UVI80_RES_RELAY(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Dim Relay_On As String
    Dim Test_PinName As String


    Relay_On = argarray(1)
    Test_PinName = argarray(0)
        
        Call Relay_Circuit_check(tNum, 0.05, 0.004, 1.7, 2.3, Relay_On, 0.003, Test_PinName, "")
        
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVI80_RES_RELAY"
End Sub



Private Sub FT_DIBC_Lib_UVS64_C(tNum As String, argarray As Variant)
On Error GoTo errHandler
    ' argArray =  Array("VDD12_ADC:4.15e-05","VDD12_AMUX:3.92e-05", ...)
    Dim tokens() As String
    Dim i        As Integer
    Dim SlotType As String
    For i = 0 To UBound(argarray)
        tokens = Split(argarray(i), ":")  ' assert UBound(tokens) = 1
'        If tokens(0) = "VDD_ECPU" Then
'            Stop
'        End If
        SlotType = GetInstrument_PWR_Pin(CStr(tokens(0)), 0)
        Call DIBC_Lib_HEXVS_C_Open_MeasureCapacitor_2(tNum, tokens(0), LCase(SlotType), Format(tokens(1), "general number"), 0.15, 1.2, -1, -1, 0)
    Next i

    tNum = CStr(CLng(tNum) + 10)
    For i = 0 To UBound(argarray)
        tokens = Split(argarray(i), ":")
        SlotType = GetInstrument_PWR_Pin(CStr(tokens(0)), 0)
        Call UVS256_HexVs_Leakage(tNum, 0.1, 0, 0.002, tokens(0), LCase(SlotType))
    Next i
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVS64_C"
End Sub
Private Sub FT_DIBC_Lib_UVS256HP_C(tNum As String, argarray As Variant)
On Error GoTo errHandler
 ' argArray =  Array("VDD12_ADC:4.15e-05","VDD12_AMUX:3.92e-05", ...)
    Dim tokens() As String
    Dim i        As Integer
    Dim SlotType As String
    For i = 0 To UBound(argarray)
        tokens = Split(argarray(i), ":")  ' assert UBound(tokens) = 1
        SlotType = GetInstrument_PWR_Pin(CStr(tokens(0)), 0)
        Call DIBC_Lib_HEXVS_C_Open_MeasureCapacitor_2(tNum, tokens(0), LCase(SlotType), Format(tokens(1), "general number"), 0.15, 1.2, -1, -1, 0)
    Next i

    tNum = CStr(CLng(tNum) + 10)
    For i = 0 To UBound(argarray)
        tokens = Split(argarray(i), ":")
        SlotType = GetInstrument_PWR_Pin(CStr(tokens(0)), 0)
        Call UVS256_HexVs_Leakage(tNum, 0.1, 0, 0.0002, tokens(0), LCase(SlotType))
    Next i
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVS256HP_C"
End Sub
Private Sub FT_DIBC_Lib_UVS256HP_R(tNum As String, argarray As Variant)
End Sub
Private Sub FT_DIBC_Lib_UP2200_R(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Dim Relay_On As String
    Dim Route_Start As String
    Dim Route_End As String
    Dim Rly_temp As String
    Dim TestPin_temp As String
    Dim TestPin_ary() As String
    Dim i As Integer
    'Route_Start = argArray(0)
    Relay_On = argarray(1)
    For i = 0 To UBound(argarray)
        Route_Start = argarray(i)
        Call Relay_Circuit_check(tNum, 0.05, 0.004, 1.7, 2.3, , 0.003, Route_Start, "")
    Next i
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP2200_R"
End Sub
Private Sub FT_DIBC_Lib_COMPLEX_PLUS(tNum As String, argarray As Variant)
On Error GoTo errHandler
  'argArray = Array("VDDQL_DDR", "DDR0_RREF", "K31", "DDR1_RREF", "K32", "DDR2_RREF", "K33", "DDR3_RREF", "K34")
    ' Dim powerPin As String: powerPin = argArray(0)
    ' Dim p1       As String
    ' Dim k1       As String

    ' Dim i As Integer
    ' For i = 1 To UBound(argArray) Step 2
        ' If i + 1 >= Ubound(argArray) Then Exit For
        ' p1 = argArray(i)
        ' k1 = argArray(i + 1)
        ' Call Relay_Circuit_check(tNum, 6, 0.004, 1.7, 2.3, k1, 0.003, powerPin, p1)
    ' Next i
    Dim powerPin As String: powerPin = argarray(0)
    Dim p1       As String
    Dim k1       As String
    Dim k2       As String
    
    
    Dim i As Integer
    For i = 1 To UBound(argarray) Step 3
        If i + 1 >= UBound(argarray) Then Exit For
        p1 = argarray(i)
        k1 = argarray(i + 1)
        k2 = argarray(i + 2)
        TheHdw.Utility.pins(k1).State = tlUtilBitOff
        Call Relay_Circuit_check(tNum, 0.05, 0.004, 1.7, 2.3, k2, 0.003, p1, powerPin)
    Next i
        
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_COMPLEX_PLUS"
End Sub
Private Sub FT_DIBC_Lib_UVS256HP_RELAY_R(tNum As String, argarray As Variant)
On Error GoTo errHandler
'argArray = Array("PAD_MTR_ANALOG_TEST_N","K10","1.20e-06")
    Dim Relay_On As String
    Dim Test_PinName As String

    Relay_On = argarray(1)
    Test_PinName = argarray(0)
    Call DCVI_Cap_meas_1(tNum, Test_PinName, Relay_On)

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVS256HP_RELAY_R"
End Sub


Private Sub FT_DIBC_Lib_UVS256_RELAY_R(tNum As String, argarray As Variant)
On Error GoTo errHandler
'argArray = Array("PAD_MTR_ANALOG_TEST_N","K10","1.20e-06")
    Dim Relay_On As String
    Dim Test_PinName As String

    Relay_On = argarray(1)
    Test_PinName = argarray(0)
    Call DCVI_Cap_meas_1(tNum, Test_PinName, Relay_On)

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVS256_RELAY_R"
End Sub


'221012 Ben Modify

Private Sub FT_DIBC_Lib_Uvi80_C_Cap(tNum As String, argarray As Variant)

        'Array("VDD_SOC_0","C1006","1.593e-05")
    On Error GoTo errHandler
        'Dim p1 As String: p1 = argArray(0)
        
        Dim Test_PinName As String
    Dim Total_Cap As Double
    Dim Cap_Current As Double
    Dim Cap_ChargeVoltage As Double
    Dim Cap_Voltageoffset As Double
    
        Test_PinName = argarray(0)
        Total_Cap = argarray(2)
        
        
        If Total_Cap <= 0.0000001 Then
    Cap_Current = 0.00005
    Cap_ChargeVoltage = 0.2
    Cap_Voltageoffset = 0.5
    ElseIf Total_Cap <= 0.000001 Then
    Cap_Current = 0.0005
    Cap_ChargeVoltage = 0.2
    Cap_Voltageoffset = 0.5
    ElseIf Total_Cap >= 0.00001 Then
    Cap_Current = 0.015
    Cap_ChargeVoltage = 0.2
    Cap_Voltageoffset = 0.5
    Else
    Cap_Current = 0.005
    Cap_ChargeVoltage = 0.2
    Cap_Voltageoffset = 0.5
    End If
        
    Call UVI80_Cap_Leakage(tNum, Test_PinName)
        Call DCVI_Cap_Meas_2(tNum, Test_PinName, "", Cap_Current, Cap_ChargeVoltage, Cap_Voltageoffset)
        
        
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_Uvi80_C_Cap"
End Sub

' Private Sub DIBC_Lib_Uvi80_C_Cap(tNum As String, argArray As Variant)


    ' On Error GoTo ErrHandler
        ' Dim p1 As String: p1 = argArray(0)
    
    ' Call UVI80_Cap_Leakage(tNum, p1)

' Exit Sub
' ErrHandler:
    ' TheExec.Datalog.WriteComment "error in DIBC_Lib_Uvi80_C_Cap"
' End Sub




Private Sub FT_DIBC_Lib_UP2200_RELAY_R(tNum As String, argarray As Variant)
On Error GoTo errHandler
' argArray = Array("ATC1_USB_RESREF","K32","K33")
    Dim p1 As String: p1 = argarray(0)
    Dim k1 As String: k1 = argarray(1)
    Dim k2 As String: k2 = argarray(2)

    ' toggle relay (k1: OFF, k2: OFF)
    TheHdw.Utility.pins(k1).State = tlUtilBitOff
    TheHdw.Utility.pins(k2).State = tlUtilBitOff

    Call DIBC_measureRelayTDRDelta_2(tNum, k1, p1)

    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), k1, p1)

    ' toggle relay (K1: OFF, k2: ON)
    TheHdw.Utility.pins(k1).State = tlUtilBitOff
    TheHdw.Utility.pins(k2).State = tlUtilBitOn

    Call Relay_Circuit_check(CStr(CLng(tNum) + 20), 0.05, 0.004, 1.7, 2.3, k1, 0.003, p1, "")

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP2200_RELAY_R"
End Sub


Private Sub FT_DIBC_Lib_UP2200_RELAY_C_3(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Dim p1 As String: p1 = argarray(0)
    Dim p2 As String: p2 = argarray(1)
    Dim k1 As String: k1 = argarray(2)
    'Dim k2 As String: k2 = argArray(2)

    ' toggle relay (k1: OFF, k2: OFF)
    TheHdw.Utility.pins(k1).State = tlUtilBitOff
    'TheHdw.Utility.pins(k2).State = tlUtilBitOff

    Call DIBC_measureRelayTDRDelta_2(tNum, k1, p1)
    Call DIBC_measureRelayTDRDelta_2(tNum, k1, p2)
    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), k1, p1)
    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), k1, p2)
    ' toggle relay (K1: OFF, k2: ON)
    TheHdw.Utility.pins(k1).State = tlUtilBitOff
    'TheHdw.Utility.pins(k2).State = tlUtilBitOn

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP2200_RELAY_C_3"
End Sub


Private Sub FT_DIBC_Lib_UP2200_RELAY_C_4(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Dim p1 As String: p1 = argarray(0)
    Dim p2 As String: p2 = argarray(1)
    Dim k1 As String: k1 = argarray(2)
    Dim k2 As String: k1 = argarray(2)
    Dim cap As String: cap = argarray(3)
    'Dim k2 As String: k2 = argArray(2)

    ' toggle relay (k1: OFF, k2: OFF)
    TheHdw.Utility.pins(k1).State = tlUtilBitOff
    'TheHdw.Utility.pins(k2).State = tlUtilBitOff

    Call DIBC_measureRelayTDRDelta_2(tNum, k1, p1)
        Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), k1, p1)
        
    Call DIBC_measureRelayTDRDelta_2(CStr(CLng(tNum) + 20), k1, p2)
    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 30), k1, p2)
        
        Call Relay_Circuit_check(CStr(CLng(tNum) + 40), 0.05, 0.004, 1.7, 2.3, k1, 0.003, p1, "")
        
        Call DIBC_Lib_Loop_Back_MCC_Cap_meas(CStr(CLng(tNum) + 50), Array(p1, p2), 0.02)
    ' toggle relay (K1: OFF, k2: ON)
    TheHdw.Utility.pins(k1).State = tlUtilBitOff
    'TheHdw.Utility.pins(k2).State = tlUtilBitOn

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP2200_RELAY_C_3"
End Sub


Private Sub FT_DIBC_Lib_UP2200_RELAY_R2(tNum As String, argarray As Variant)
On Error GoTo errHandler
     ' argArray = Array("RT_CLK32768","RT_CLK32768_PA","K01")
    Dim RLY_name As String: RLY_name = argarray(1)
    Dim Test_Pin As String: Test_Pin = argarray(0)

    Call DIBC_measureRelayTDRDelta_2(tNum, RLY_name, Test_Pin)
    Call DIBC_PPMU_Relay_FIMV_rising_new(CStr(CLng(tNum) + 10), RLY_name, Test_Pin)
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UP2200_RELAY_R2"
End Sub
Private Sub FT_DIBC_Lib_UVS256HP_RELAY_R1(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Dim Relay_On As String
    Dim Route_Start As String
    Dim Route_End As String
    Dim Rly_temp As String
    Dim TestPin_temp As String
    Dim TestPin_ary() As String
    Dim i As Integer
    Route_Start = argarray(0)
    Relay_On = argarray(1)
    Call Relay_Circuit_check(tNum, 0.05, 0.004, 1.7, 2.3, Relay_On, 0.003, Route_Start, "")
     
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_UVS256HP_RELAY_R1"
End Sub





Private Function Read_Status_Reg(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Dim Relay_On As String
    Relay_On = Join(argarray, ",")
    ' connect relay
   ' TheHdw.Utility.pins("K02").State = tlUtilBitOn
   ' TheHdw.Utility.pins("K04").State = tlUtilBitOn
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn

    ' level timing pattern
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call SetupHRam
    Dim spiRomPat As String: spiRomPat = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\SPIROMReadStatus.PAT"
    If dibonTesterType = "Jaguar" Then
        TheHdw.Digital.patterns.pat(spiRomPat).Load
    ElseIf dibonTesterType = "UltraFLEXplus" Then
        TheHdw.patterns(spiRomPat).Load
    End If
    ' read register
    get_status_reg2 &H5, 8

    ' disconnect relay
    'TheHdw.Utility.pins("K02").State = tlUtilBitOff
    'TheHdw.Utility.pins("K04").State = tlUtilBitOff
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Read_Status_Reg"
End Function

Private Function get_status_reg2(data As Integer, data_len As Integer) As String
On Error GoTo errHandler
    get_status_reg2 = ""

    Dim pat_modify() As String
    Dim numcap As Long
    Dim PinData As New PinListData
    Dim regdata() As Long
    Dim rawdata As String
    Dim modify_data As String: modify_data = ""
    ReDim pat_modify(data_len - 1)
    
    Call get_data(data, pat_modify, data_len)

    Dim i As Integer
    For i = 0 To data_len - 1
        modify_data = modify_data + pat_modify(i)
    Next i
        Dim spiRomPat As String: spiRomPat = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\SPIROMReadStatus.PAT"

    Call TheHdw.Digital.pins("SPI1_MOSI").patterns(spiRomPat).ModifyVectorBlockData("", 4, pat_modify)
    TheHdw.Wait 0.05

    TheHdw.patterns(spiRomPat).start
    TheHdw.Digital.Patgen.HaltWait
    numcap = 8
    PinData = TheHdw.Digital.pins("SPI1_MISO").hram.PinPF(0, , numcap)

    Dim site As Variant
    Dim sdbRegValue1 As New SiteDouble
    Dim sdbRegValue2 As New SiteDouble
    Dim sdbRegValue3 As New SiteDouble
    Dim sdbRegValue4 As New SiteDouble
    Dim sdbRegValue5 As New SiteDouble
    Dim pldMeasureValue As New PinListData: pldMeasureValue.AddPin "SPI1_MISO"
    Dim ReadStatusReg As Integer
    Dim ReadFlagStatusReg As Integer
    Dim ReadEnhancedVolatileConfReg As Integer
    Dim ReadExtendedAddressReg As Integer
    Dim ReadVolatileConfReg As Integer

    For Each site In TheExec.sites.Active
        regdata = PinData.pins(0).value

        For i = 0 To numcap - 1
            If i Mod 4 = 0 Then get_status_reg2 = get_status_reg2 + " "

            If regdata(i) = 2 Then
                get_status_reg2 = get_status_reg2 + "1"
            Else
                get_status_reg2 = get_status_reg2 + "0"
            End If
        Next i

        If (data = &H5) Then
            ReadStatusReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
            sdbRegValue1.value = ReadStatusReg
        ElseIf (data = &H70) Then
            ReadFlagStatusReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
            sdbRegValue2.value = ReadFlagStatusReg
        ElseIf (data = &H65) Then
            ReadEnhancedVolatileConfReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
            sdbRegValue3.value = ReadEnhancedVolatileConfReg
        ElseIf (data = &HC8) Then
            ReadExtendedAddressReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
            sdbRegValue4.value = ReadExtendedAddressReg
        ElseIf (data = &H85) Then
            ReadVolatileConfReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
            sdbRegValue5.value = ReadVolatileConfReg
        End If

    Next site

    If data = &H5 Then
        pldMeasureValue.pins("SPI1_MISO") = sdbRegValue1
        TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=0, hiVal:=0, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:="ReadStatusReg", ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
    End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in get_status_reg2"
End Function

Private Function Read_Device_ID(tNum As String, argarray As Variant)
On Error GoTo errHandler
    Dim i As Integer
    Dim pat_modify(7) As String
    Dim numcap As Long
    Dim PinData As New PinListData
    Dim regdata() As Long
    Dim rawdata As String: rawdata = ""
    Dim modify_data As String: modify_data = ""
    Dim Relay_On As String
    Relay_On = Join(argarray, ",")
    ' connect relay
    'TheHdw.Utility.pins("K02").State = tlUtilBitOn
    'TheHdw.Utility.pins("K04").State = tlUtilBitOn
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
    ' set DCVS
    With TheHdw.DCVS.pins("QSPI_PWR_1P2")
        .Connect tlDCVSConnectDefault
        TheHdw.Wait 0.0005
        .mode = tlDCVSModeVoltage
        TheHdw.Wait 0.01
        .Voltage.Main.value = 1.2 'AL20181114, 1.8v will cause the test result not stable.
        .Voltage.Alt.value = 1.8
        .Voltage.Output = tlDCVSVoltageMain
        .CurrentRange.value = 0.1
        .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        .CurrentLimit.Source.FoldLimit.level.value = 0.1
        .CurrentLimit.Source.FoldLimit.TimeOut.value = 0.05
        .Meter.mode = tlDCVSMeterVoltage
        '.Meter.VoltageRange = 18
        .Meter.Filter.bypass = False
        .Alarm(tlDCVSAlarmAll) = tlAlarmDefault
        .Gate = True
        TheHdw.Wait 0.01
    End With
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    If dibonTesterType = "Jaguar" Then  'uflx 20201102
        With TheHdw.PinLevels.pins("SPI0_MISO")
            .ModifyLevel chVil, -0.15
            .ModifyLevel chVih, 1.8
            .ModifyLevel chVol, 0.8
            .ModifyLevel chVoh, 1
            .ModifyLevel chIol, 0.001
            .ModifyLevel chIoh, -0.001
            .ModifyLevel chVt, 0.9
            TheHdw.Wait 0.05
        End With
    ElseIf dibonTesterType = "UltraFLEXplus" Then
        TheHdw.Digital.pins("SPI0_MISO").Connect
        TheHdw.Wait 0.002
        TheHdw.Digital.pins("SPI0_MISO").Levels.value(chVih) = 1.8
        TheHdw.Digital.pins("SPI0_MISO").Levels.value(chVil) = -0.15
        TheHdw.Digital.pins("SPI0_MISO").Levels.value(chVol) = 0.8
        TheHdw.Digital.pins("SPI0_MISO").Levels.value(chVch) = 5.5
        TheHdw.Digital.pins("SPI0_MISO").Levels.value(chVcl) = -1
        TheHdw.Digital.pins("SPI0_MISO").Levels.value(chVoh) = 1
        'TheHdw.Digital.pins("SPI0_MISO").Levels.value(chIoh) = 0.001
        'TheHdw.Digital.pins("SPI0_MISO").Levels.value(chIol) = -0.001
        TheHdw.Digital.pins("SPI0_MISO").Levels.value(chVt) = 0.9

        TheHdw.Digital.pins("SPI0_MISO").Connect   'uflx 20201102
    End If
    Call SetupHRam
    Dim spiRomDebugPat As String: spiRomDebugPat = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\SPIROMReadID_Debug.PAT"
    If dibonTesterType = "Jaguar" Then
        TheHdw.Digital.patterns.pat(spiRomDebugPat).Load
    ElseIf dibonTesterType = "UltraFLEXplus" Then
        TheHdw.patterns(spiRomDebugPat).Load
    End If
    'read ID
    get_status_reg &H9F, 8, tNum

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Read_Device_ID"
End Function

Private Function get_status_reg(data As Integer, data_len As Integer, tNum As String) As String
On Error GoTo errHandler
    get_status_reg = ""

    Dim i As Integer
    Dim pat_modify() As String
    Dim numcap As Long
    Dim PinData As New PinListData
    Dim regdata() As Long
    Dim rawdata As String
    Dim modify_data As String: modify_data = ""
    ReDim pat_modify(data_len - 1)

    Call get_data(data, pat_modify, data_len)
    For i = 0 To data_len - 1
        modify_data = modify_data + pat_modify(i)
    Next i

    Dim spiRomDebugPat As String: spiRomDebugPat = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\SPIROMReadID_Debug.PAT"

    Call TheHdw.Digital.pins("SPI1_MOSI").patterns(spiRomDebugPat).ModifyVectorBlockData("", 4, pat_modify)
    TheHdw.Wait 0.05
    TheHdw.patterns(spiRomDebugPat).start
    TheHdw.Digital.Patgen.HaltWait
    numcap = 160
    PinData = TheHdw.Digital.pins("SPI1_MISO").hram.PinPF(0, , numcap)

    Dim site As Variant
    Dim sdbRegValue1 As New SiteDouble
    Dim sdbRegValue2 As New SiteDouble
    Dim sdbRegValue3 As New SiteDouble
    Dim pldMeasureValue As New PinListData: pldMeasureValue.AddPin "SPI1_MISO"
    Dim status_reg_ManufacturerID As Integer
    Dim status_reg_MemoryType As Integer
    Dim status_reg_MemoryCapacity As Integer

    For Each site In TheExec.sites.Active
        regdata = PinData.pins(0).value

        For i = 0 To numcap - 1
            If i Mod 4 = 0 Then get_status_reg = get_status_reg + " "

            If regdata(i) = 2 Then
            get_status_reg = get_status_reg + "1"
            Else
            get_status_reg = get_status_reg + "0"
            End If
        Next i

        '20H
        status_reg_ManufacturerID = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
        sdbRegValue1.value = status_reg_ManufacturerID
        'BBH
        status_reg_MemoryType = (val(regdata(8)) - 1) * 2 ^ 7 + (val(regdata(9)) - 1) * 2 ^ 6 + (val(regdata(10)) - 1) * 2 ^ 5 + (val(regdata(11)) - 1) * 2 ^ 4 + (val(regdata(12)) - 1) * 2 ^ 3 + (val(regdata(13)) - 1) * 2 ^ 2 + (val(regdata(14)) - 1) * 2 ^ 1 + (val(regdata(15)) - 1) * 2 ^ 0
        sdbRegValue2.value = status_reg_MemoryType
        '19H
        status_reg_MemoryCapacity = (val(regdata(16)) - 1) * 2 ^ 7 + (val(regdata(17)) - 1) * 2 ^ 6 + (val(regdata(18)) - 1) * 2 ^ 5 + (val(regdata(19)) - 1) * 2 ^ 4 + (val(regdata(20)) - 1) * 2 ^ 3 + (val(regdata(21)) - 1) * 2 ^ 2 + (val(regdata(22)) - 1) * 2 ^ 1 + (val(regdata(23)) - 1) * 2 ^ 0
        sdbRegValue3.value = status_reg_MemoryCapacity

    Next site

    pldMeasureValue.pins("SPI1_MISO") = sdbRegValue1
    Util_Judge_Result pldMeasureValue, tNum, "ManufacturerID", ""
    'TheExec.Flow.TestLimit resultVal:=pldMeasureValue, lowVal:=32, hiVal:=32, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:="ManufacturerID", forceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone

    pldMeasureValue.pins("SPI1_MISO") = sdbRegValue2
    Util_Judge_Result pldMeasureValue, CStr(CLng(tNum) + 1), "MemoryType", ""
    'TheExec.Flow.TestLimit resultVal:=pldMeasureValue, lowVal:=187, hiVal:=187, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:="MemoryType", forceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone

    pldMeasureValue.pins("SPI1_MISO") = sdbRegValue3
    Util_Judge_Result pldMeasureValue, CStr(CLng(tNum) + 2), "MemoryCapacity", ""
    'TheExec.Flow.TestLimit resultVal:=pldMeasureValue, lowVal:=25, hiVal:=25, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:="MemoryCapacity", forceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
    
    TheHdw.Utility.pins("K57").State = tlUtilBitOff
    TheHdw.Utility.pins("K58").State = tlUtilBitOff

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in get_status_reg"
End Function

Private Sub get_data(data As Integer, ByRef pat_modify() As String, data_len As Integer)
On Error GoTo errHandler
    Dim i As Integer
    For i = 0 To data_len - 1
        If (data And 2 ^ (data_len - 1 - i)) <> 0 Then
            pat_modify(i) = "1"
        Else
            pat_modify(i) = "0"
        End If
    Next i
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in get_data"
End Sub

Private Function SetupHRam()
On Error GoTo errHandler
    With TheHdw.Digital
        .Patgen.Events.Clear
        .Patgen.Events.SetCycleCount True, 0, tlCycleTypeAbsolute
        .hram.CaptureType = captSTV
        .hram.SetTrigger trigFirst, False, 0
        .Patgen.HaltMode = 1
        .hram.size = 512
    End With
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in SetupHRam"
End Function

Private Function FreeRunclk_Enable_dibchecker(tNum As String, PortName As String, DUTPin As String, Bufferpin As String, Tname As String)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "FreeRunclk_Enable_dibchecker"
    
'    On Error GoTo errHandler
    
    Dim site As Variant
    Dim i As Long
    Dim PLLLockChecked As New SiteLong
    Dim measf As New PinListData
    Dim NotLocked As Boolean
    Dim XI0_REFCLK As String
    Dim PortMode As String
    Dim PLL_Lock As New SiteLong
    Dim FreeRunPortList() As String
    Dim PortNum As Integer
    
    FreeRunPortList = Split(PortName, ",")
    For Each site In TheExec.sites.Active
        PLLLockChecked = 1
        PLL_Lock = 1
    Next site
    
    TheHdw.Protocol.ports(PortName).Halt
    TheHdw.Protocol.ports(PortName).Enabled = False
        
    TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, PinLevelsSheet:="DIBC_TSets__FT", TimeSetSheet:="DIBC_Levels__FT" ''May2 added
        '''0831 by rita
        TheHdw.Digital.pins(PortName).Levels.value(chVil) = 0  'm_sAllDigitalPinList
        TheHdw.Digital.pins(PortName).Levels.value(chVih) = 1.8
        TheHdw.Digital.pins(PortName).Levels.value(chIol) = 0
        TheHdw.Digital.pins(PortName).Levels.value(chIoh) = 0
        TheHdw.Digital.pins(PortName).Levels.value(chVch) = 6
        TheHdw.Digital.pins(PortName).Levels.value(chVoutLoTyp) = 0
        TheHdw.Digital.pins(PortName).Levels.value(chVoutHiTyp) = 0
        TheHdw.Digital.pins(PortName).Levels.DriverMode = tlDriverModeLargeHiZ  'Largeswing-HiZ
        
        TheHdw.Digital.pins(DUTPin).Levels.value(chVol) = 0.54
        TheHdw.Digital.pins(DUTPin).Levels.value(chVoh) = 0.54
        TheHdw.Digital.pins(DUTPin).Levels.value(chVt) = 0.9
        TheHdw.Digital.pins(DUTPin).Levels.value(chVcl) = 0
        
        TheHdw.Digital.pins(Bufferpin).Levels.value(chVol) = 0.55
        TheHdw.Digital.pins(Bufferpin).Levels.value(chVoh) = 0.55
        TheHdw.Digital.pins(Bufferpin).Levels.value(chVt) = 0.6
        TheHdw.Digital.pins(Bufferpin).Levels.value(chVcl) = -1
            ''''
        TheHdw.Wait 0.005 '0904 by rita
        TheHdw.Protocol.ports(PortName).Enabled = True
        TheHdw.Protocol.ports(PortName).NWire.ResetPLL
        TheHdw.Wait 0.001

        ' check PLL lock status
        For Each site In TheExec.sites.Active
            For PortNum = 0 To UBound(FreeRunPortList)
                If TheHdw.Protocol.ports(FreeRunPortList(PortNum)).NWire.IsPLLLocked = False Then
                    PLL_Lock = 0
                End If
            Next
        Next site
        
        ' start nWire engine.
        Call TheHdw.Protocol.ports(PortName).NWire.Frames("RunFreeClock").Execute
        TheHdw.Protocol.ports(PortName).IdleWait
        TheHdw.Wait 0.001
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites.Selected
                PLL_Lock(site) = 1
            Next site
        End If
        
        Dim pldMeasureValue As New PinListData
        pldMeasureValue.AddPin PortName
        pldMeasureValue.pins(PortName) = PLL_Lock
        Util_Judge_Result pldMeasureValue, tNum, "IC", ""

'    Exit Function
    

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function


Private Function FreeRunClk_Disable_dibchecker(PortName As String) As Long
On Error GoTo errHandler
    Dim site As Variant
    TheHdw.Protocol.ports(PortName).Halt
    TheHdw.Protocol.ports(PortName).Enabled = False     'scope out point
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in FreeRunClk_Disable_dibchecker"
End Function


Private Function MeasFreq_dibchecker(tNum As String, mea_pin As String, period As Double, Tname As String, Optional collectdata As Boolean = False) As Long ', Tnum As Long, Tname As String, Optional collectdata As Boolean = False) As Long
On Error GoTo errHandler
    Dim freqCnt1 As New PinListData
    Dim MeasFreq1 As New SiteDouble
    Dim TimeInterval1 As New SiteDouble
    Dim StartVolt As Double
    Dim stepvolt As Double
    Dim stepcount As Integer
    Dim PinNum As Integer
    Dim i As Integer
    Dim currvolt As Double
    Dim site As Variant
    
    StartVolt = 0.1
    stepvolt = 0.05
    stepcount = 30
    If (collectdata = True) Then
        For i = 0 To stepcount
        
        currvolt = StartVolt + i * stepvolt
    
            TheHdw.Digital.pins(mea_pin).Levels.value(chVoh) = currvolt
            With TheHdw.Digital.pins(mea_pin).FreqCtr
                .EventSource = VOH
                .EventSlope = Positive
                .Enable = IntervalEnable
                .Interval = period
            End With
            
            TheHdw.Digital.pins(mea_pin).FreqCtr.Clear
            TheHdw.Digital.pins(mea_pin).FreqCtr.start
            freqCnt1 = TheHdw.Digital.pins(mea_pin).FreqCtr.Read
            TimeInterval1 = TheHdw.Digital.pins(mea_pin).FreqCtr.Interval
            MeasFreq1 = freqCnt1.Math.divide(TimeInterval1)
            Debug.Print ("measPin: " + mea_pin + " Voh=" + CStr(TheHdw.Digital.pins(mea_pin).Levels.value(chVoh)) + " Frequcney=" + CStr(MeasFreq1(0) / 1000000#) + " MHz")
            TheExec.Datalog.WriteComment ("measPin: " + mea_pin + " Voh=" + CStr(TheHdw.Digital.pins(mea_pin).Levels.value(chVoh)) + " Frequency=" + CStr(MeasFreq1(0) / 1000000#) + " MHz")
      
                 
        Next i
    Else
        With TheHdw.Digital.pins(mea_pin).FreqCtr
            .EventSource = VOH
            .EventSlope = Positive
            .Enable = IntervalEnable
            .Interval = period
        End With
        
        TheHdw.Digital.pins(mea_pin).FreqCtr.Clear
        TheHdw.Digital.pins(mea_pin).FreqCtr.start
        freqCnt1 = TheHdw.Digital.pins(mea_pin).FreqCtr.Read
       For PinNum = 0 To freqCnt1.pins.Count - 1
            TimeInterval1 = TheHdw.Digital.pins(freqCnt1.pins(PinNum).name).FreqCtr.Interval
            MeasFreq1 = freqCnt1.pins(PinNum).divide(TimeInterval1)
            For Each site In TheExec.sites.Active
'              theexec.Datalog.WriteComment ("Site: " + CStr(site) + "  measPin: " + freqCnt1.Pins(pinnum).Name + " Frequcney=" + CStr(MeasFreq1(site) / 1000000#) + " MHz")
'              Debug.Print ("Site: " + CStr(site) + "  measPin: " + freqCnt1.Pins(pinnum).Name + " Frequcney=" + CStr(MeasFreq1(site) / 1000000#) + " MHz")
                                       
            Next site
            
          
            Dim pldMeasureValue As New PinListData
            pldMeasureValue.AddPin freqCnt1.pins(PinNum).name
            pldMeasureValue.pins(freqCnt1.pins(PinNum).name) = MeasFreq1
            
           ' judgeCheckResult pldMeasureValue, Tnum, Tname
           'TheExec.Flow.TestLimit resultVal:=pldMeasureValue, PinName:=PinNum, lowVal:=39996000, hiVal:=40000800, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="Hz", ForceResults:=tlForceNone
          ' TheExec.Flow.TestLimit resultVal:=pldMeasureValue, PinName:=PinNum, lowVal:=23996000, hiVal:=24100800, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="Hz", ForceResults:=tlForceNone
           Util_Judge_Result pldMeasureValue, tNum, "FRC Freq", "Hz"
             
        Next PinNum
    
    End If
   
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in MeasFreq_dibchecker"
End Function

Private Function Util_Meas_Freq(MeasPin As String, period As Double) As PinListData  ' only support single measPin
On Error GoTo errHandler
    ' set freq counter
    With TheHdw.Digital.pins(MeasPin).FreqCtr
        .EventSource = VOH
        .EventSlope = Positive
        .Enable = IntervalEnable
        .Interval = period
    End With
    
    ' meas count
    Dim cnt As New PinListData
    If TheExec.TesterMode <> testModeOffline Then
        TheHdw.Digital.pins(MeasPin).FreqCtr.Clear
        TheHdw.Digital.pins(MeasPin).FreqCtr.start
        cnt = TheHdw.Digital.pins(MeasPin).FreqCtr.Read
    End If
    
    ' convert count to freq and update result
    Dim meas As New PinListData
    Dim freq As New SiteDouble
    Dim i    As Integer
    For i = 0 To cnt.pins.Count - 1
        freq = cnt.pins(i).divide(period)
        meas.AddPin MeasPin
        meas.pins(MeasPin) = freq
    Next i
           
    Set Util_Meas_Freq = meas
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Util_Meas_Freq"
End Function


Private Function DIBC_PPMU_Relay_FIMV_rising_new(TestNum As String, RelayName As Variant, TestPinName As String)
On Error GoTo errHandler
    Dim pinname_str As String:  pinname_str = TestPinName
    Dim PPMUMeasure As New PinListData
    Dim PPMUMeas() As New SiteDouble
    Dim TimeMeas() As New SiteDouble
    Dim SWTime As New SiteDouble
    Dim i As Integer
    Dim j As Integer
    Dim site As Variant
    Dim RStep As Integer
    Dim RStep_2 As Integer
    Dim Start_time As Double
    Dim End_time As Double
    Dim Swtich_DeltaT As New PinListData

    For Each site In TheExec.sites.Selected
        SWTime(site) = 0
    Next site

    ReDim PPMUMeas(100) As New SiteDouble
    ReDim TimeMeas(100) As New SiteDouble
    RStep = 20
    RStep_2 = 100
        If (pinname_str = "SPI1_SSIN") Or (pinname_str = "SPI1_SCLK") Or (pinname_str = "SPI1_MISO") Or (pinname_str = "SPI1_MOSI") Then
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOn
        Else
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOff
        End If

      TheHdw.Digital.pins(pinname_str).Disconnect

    If (pinname_str = "RT_CLK32768") Or (pinname_str = "XI0") Or (pinname_str = "XO0") Then
        TheHdw.Digital.pins(pinname_str & "_PA").Disconnect
    End If

  For Each site In TheExec.sites.Selected


    If (pinname_str = "RT_CLK32768") Or (pinname_str = "XI0") Or (pinname_str = "XO0") Then
        With TheHdw.PPMU.pins(pinname_str & "_PA")
            .Gate = tlOff
            .Connect
            .Gate = tlOn
            TheHdw.Wait 0.001
            .ForceV 0, 0
        End With '20200416
    End If

    With TheHdw.PPMU.pins(pinname_str)
        .Gate = tlOff
        .Connect
        .Gate = tlOn
        TheHdw.Wait 0.001
        .ForceV 0, 0
        TheHdw.Wait 0.01
        .ForceI 0.000002, 0.00002
        .ClampVHi = 3
    End With '20200407

        For i = 0 To RStep
            Start_time = TheExec.Timer(0)
            PPMUMeas(i) = TheHdw.PPMU.pins(pinname_str).Read(tlPPMUReadMeasurements, 2)   'normal measure
            TimeMeas(i) = TheExec.Timer(Start_time)
        Next i

        If (pinname_str = "SPI1_SSIN") Or (pinname_str = "SPI1_SCLK") Or (pinname_str = "SPI1_MISO") Or (pinname_str = "SPI1_MOSI") Then
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOff
        Else
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOn
        End If

        For i = RStep + 1 To RStep_2
            Start_time = TheExec.Timer(0)
            PPMUMeas(i) = TheHdw.PPMU.pins(pinname_str).Read(tlPPMUReadMeasurements, 2)   'normal measure
            TimeMeas(i) = TheExec.Timer(Start_time)
        Next i

        If (pinname_str = "SPI1_SSIN") Or (pinname_str = "SPI1_SCLK") Or (pinname_str = "SPI1_MISO") Or (pinname_str = "SPI1_MOSI") Then
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOn
        Else
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOff
        End If

  Next site

     Dim Catch_index As New SiteLong
    For Each site In TheExec.sites.Selected
      For i = RStep + 1 To RStep_2
        If PPMUMeas(i).Subtract(PPMUMeas(i - 1)) < -0.01 Or PPMUMeas(i).Subtract(PPMUMeas(i - 1)) > 0.07 Then
            Catch_index = i
            Exit For
        End If
      Next i
    Next site

    For Each site In TheExec.sites.Selected
      For j = RStep + 1 To Catch_index(site)
          SWTime = SWTime + TimeMeas(j)
      Next j
    Next site
    Swtich_DeltaT.AddPin (pinname_str)
    Swtich_DeltaT = SWTime

    Util_Judge_Result Swtich_DeltaT, TestNum, "Relay Switch Time", "S"
    TheHdw.Utility.pins(RelayName).State = tlUtilBitOff
    With TheHdw.PPMU.pins(pinname_str)
        .ForceV 0#
        .Connect
        .Gate = tlOn
    End With

    If (pinname_str = "RT_CLK32768") Or (pinname_str = "XI0") Or (pinname_str = "XO0") Then
        With TheHdw.PPMU.pins(pinname_str & "_PA")
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With '20200416
    End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_PPMU_Relay_FIMV_rising_new"
End Function

' Measure the delta of rising point when the relay ON and Off, the delta will be same as TDR difference
Private Function DIBC_measureRelayTDRDelta_2(tNum As String, RelayName As Variant, PinName As String)
On Error GoTo errHandler
    'setup default parameters
    Dim m_sAllDigitalPinList As String
    Dim RelayOffTime As New SiteDouble
    Dim RelayOnTime As New SiteDouble
    Dim RelayDeltaTime As New SiteDouble
    Dim sPinName As String
    Dim sRelayName As Variant
    sPinName = PinName
    sRelayName = RelayName
    TheHdw.PPMU.pins(PinName).Disconnect
    TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, PinLevelsSheet:="DIBC_Levels__FT", TimeSetSheet:="DIBC_TSets__FT"
    
    Dim strAllDigitalPins()  As String
    Dim lngPnum As Long
    If m_sAllDigitalPinList = vbNullString Then
        Call TheExec.DataManager.GetPinNames(strAllDigitalPins, chIO, lngPnum)
        m_sAllDigitalPinList = Join(strAllDigitalPins, ",")
    End If

    If m_sAllDigitalPinList <> vbNullString Then
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVil) = 0.1
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVih) = 1.1
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVol) = 0
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVoh) = 0
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVt) = 0.1
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVcl) = -1#
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVch) = 6
    End If

    Dim tdrAllPat As String: tdrAllPat = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\DIBC_TDR_ALL__.PAT"
    TheHdw.patterns(tdrAllPat).Load
    TheHdw.patterns(tdrAllPat).start

    TheHdw.Digital.pins(sPinName).Connect
    TheHdw.Wait 0.001

    ' Set the VOh value
    TheHdw.Digital.pins(sPinName).Levels.value(chVoh) = 0.8

    If sPinName = "DDR2_RREF" Then ' for t4402 Kevin0601
        TheHdw.Digital.pins(sPinName).Levels.value(chVoh) = 0.79 '0.85 Kevin0601,0.8 Kevin0823
    ElseIf sPinName = "DDR0_RREF" Or sPinName = "DDR3_RREF" Then ' for t4400,4403 Kevin0601
        TheHdw.Digital.pins(sPinName).Levels.value(chVoh) = 0.75
    End If

    TheHdw.Wait 0.001
    ' Find the time with relay Off
    TheHdw.Utility.pins(sRelayName).State = tlUtilBitOff
    RelayOffTime = searchTDRRisingTime(sPinName)

    ' Find the time with relay on
    TheHdw.Utility.pins(sRelayName).State = tlUtilBitOn

    RelayOnTime = searchTDRRisingTime(sPinName)
    RelayDeltaTime = RelayOnTime.Subtract(RelayOffTime)

    Dim plDeltaValue As New PinListData
    Dim thissite As Variant
    plDeltaValue.AddPin (sPinName)
    For Each thissite In TheExec.sites.Selected
        plDeltaValue.pins(sPinName).value = RelayDeltaTime.Abs
    Next thissite

    Util_Judge_Result plDeltaValue, tNum, "Relay TDR Time", "S"

    TheHdw.Utility.pins(sRelayName).State = tlUtilBitOff
    TheHdw.Digital.pins(sPinName).Disconnect
    TheHdw.Wait 0.001
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_measureRelayTDRDelta_2"
End Function

Private Function searchTDRRisingTime(sPinName As String) As SiteDouble
On Error GoTo errHandler
    Dim StartTime As Double:     StartTime = 20 * ns
    Dim StopTime As Double:  StopTime = 80 * ns
    Dim StepSize As Double:     StepSize = 0.1 * ns
    Dim sdRisingTime As New SiteDouble
    Dim timeval As Double: timeval = StartTime
    Dim plFailCount As PinListData
    Dim thissite As Variant
    Dim foundtime As New SiteBoolean
              
    For Each thissite In TheExec.sites.Selected
        foundtime = False
        sdRisingTime = 0
    Next thissite

    Do While timeval < StopTime
        TheHdw.Digital.pins(sPinName).Timing.EdgeTime("DIBC_TDR__", chEdgeR0) = timeval
        ' Run the pattern
        TheHdw.Digital.Patgen.Restart
        TheHdw.Digital.Patgen.HaltWait
        Set plFailCount = TheHdw.Digital.pins(sPinName).FailCount
        For Each thissite In TheExec.sites.Selected
            ' Read the failcount
            If Not foundtime Then
                If plFailCount = 0 Then
                    foundtime = True
                    sdRisingTime = timeval
                End If
            End If
        Next thissite
        ' Found for all sites?
        If foundtime.All(True) Then
            Exit Do
        End If
        timeval = timeval + StepSize
    Loop

    Set searchTDRRisingTime = sdRisingTime

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in searchTDRRisingTime"
End Function

Private Function GetInstrument_PWR_Pin(PinList As String, site As Variant) As String
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "GetInstrument_PWR_Pin"
    Dim chanString As String
    Dim PinName() As String
    Dim NumberPins As Long
    Call TheExec.DataManager.DecomposePinList(PinList, PinName(), NumberPins)
    Call TheExec.DataManager.GetChannelStringFromPinAndSite(PinName(0), site, chanString)
    Dim slotstr() As String
    Dim slot As Long

        If chanString = "" Then
            TheExec.Datalog.WriteComment ("Warnning : Please check pin type of  " & PinList & " in channel map")
        Else
            slotstr = Split(chanString, ".")
            slot = CLng(slotstr(0))
            GetInstrument_PWR_Pin = TheHdw.config.Slots(slot).type
        End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function DCVI_Cap_meas_1(tNum As String, Test_PinName As String, Optional Relay_On As String)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DCVI_Cap_meas_1"

       ' Call SmartRelaySwitch("")

        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
        '***connect DCVI instrument***'
        TheHdw.DCVI.pins(Test_PinName).Connect tlDCVIConnectDefault
        TheHdw.DCVI.pins(Test_PinName).Disconnect tlDCVIConnectHighForce
        TheHdw.DCVI.pins(Test_PinName).LocalKelvin = True
        TheHdw.Wait 0.0005
        '***-----------------------***'

        '********************capacitor test begin********************'
        Dim vsite As Variant
        Dim intI As Integer
        Dim intJ As Integer
        Dim pldMeasureValue As New PinListData
        Dim v1 As New PinListData
        Dim v2 As New PinListData
        Dim deltaV As New PinListData
        Dim t1 As New PinListData
        Dim t2 As New PinListData
        Dim deltaT As New PinListData
        Dim dT As Double
        Dim DCVIResultDSP As New DSPWave
        Dim strDCVI_PinArray() As String
        Dim dblCapSampleRate As Double
        Dim dblCapSampleSize As Double
        Dim dblCurrentRange As Double
        Dim dblCaptureWait As Double
        Dim strTestDCVI_Pin As String
        Dim dblTestMeasureCurrent As Double
        Dim dblTestChargeVoltage As Double
        Dim dblTestOffsetVoltage As Double
        Dim iIndex1 As New SiteLong
        Dim iIndex2 As New SiteLong

        strTestDCVI_Pin = Test_PinName
        dblTestMeasureCurrent = 0.0005
        dblTestChargeVoltage = 1.5
        dblTestOffsetVoltage = 0.8

        strDCVI_PinArray = Split(strTestDCVI_Pin, ",")

        dblCapSampleRate = 100 * kHz
        dblCapSampleSize = 512
        dblCaptureWait = dblCapSampleSize / dblCapSampleRate
        dblCurrentRange = 0.005

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmOff

        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.Add ("CapSignal")
        With TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals("CapSignal")
                .SampleRate = 100000
                .SampleSize = 512
                .LoadSettings
        End With

        '***set DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Current = 0.005
                .Voltage = 5
                .VoltageRange.value = 7
                .CurrentRange.Autorange = True
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 7
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'

        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '*****measure parasitic capacitor begin*****'

        Dim dblMeasureCurrentParasiticCap As Double
        Dim dblCurrentRangeParasiticCap As Double
        Dim pldMeasureValueParasitic As New PinListData
        Dim pldV1_parasitic As New PinListData
        Dim pldV2_parasitic As New PinListData
        Dim pldDeltaV_parasitic As New PinListData
        Dim pldT1_parasitic As New PinListData
        Dim pldT2_parasitic As New PinListData
        Dim pldDeltaT_parasitic As New PinListData
        Dim pldDT_parasitic As Double
        Dim DCVIResultDSP_parasitic As New DSPWave

        TheHdw.Wait 0.003

        dblMeasureCurrentParasiticCap = 0.000001
        dblCurrentRangeParasiticCap = 0.000001

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblMeasureCurrentParasiticCap
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait 0.00512
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'computing of the parasitic cap value
        For intJ = 0 To UBound(strDCVI_PinArray)
                pldV1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldV2_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT2_parasitic.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
'''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
'''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
'''''                                        Call DCVIResultDSP_parasitic.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
'''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP_parasitic = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

        'DCVIResultDSP_parasitic.Plot

                                iIndex1 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                pldV1_parasitic.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic.data(iIndex1)
                                pldV2_parasitic.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic.data(iIndex2)
                                pldDT_parasitic = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value - pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value
                                If pldDT_parasitic = 0 Then
                                        pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        pldDeltaV_parasitic.AddPin (strTestDCVI_Pin)
        pldDeltaV_parasitic = pldV2_parasitic.Math.Subtract(pldV1_parasitic)
        pldDeltaT_parasitic.AddPin (strTestDCVI_Pin)
        pldDeltaT_parasitic = pldT2_parasitic.Math.Subtract(pldT1_parasitic)

        pldMeasureValueParasitic.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValueParasitic = pldDeltaT_parasitic.Math.Multiply(dblMeasureCurrentParasiticCap)
        pldMeasureValueParasitic = pldMeasureValueParasitic.Math.divide(pldDeltaV_parasitic)

        For intJ = 0 To UBound(strDCVI_PinArray)
                If TheExec.TesterMode = testModeOffline Then
                        For Each vsite In TheExec.sites.Active
                                pldMeasureValueParasitic.pins(strDCVI_PinArray(intJ)).value = 0
                        Next vsite
                End If
        Next intJ

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.002
       ' Call SmartRelaySwitch("K96")
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
        TheHdw.Wait 0.003

'        TheHdw.DCVI.Pins("PAD_MTR_ANALOG_TEST_N").Disconnect tlDCVIConnectHighForce
        TheHdw.DCVI.pins(strTestDCVI_Pin).LocalKelvin = False
        TheHdw.DCVI.pins(strTestDCVI_Pin).Connect tlDCVIConnectDefault
        TheHdw.Wait 0.0005

        '*****measure parasitic capacitor end*****'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblTestMeasureCurrent
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait dblCaptureWait
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False

        'computing of the capacitor value
        For intJ = 0 To UBound(strDCVI_PinArray)
                v1.AddPin (strDCVI_PinArray(intJ))
                v2.AddPin (strDCVI_PinArray(intJ))
                t1.AddPin (strDCVI_PinArray(intJ))
                t2.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
''''                                        Call DCVIResultDSP.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

        'DCVIResultDSP.Plot

                                iIndex1 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                t1.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                t2.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                v1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(iIndex1)
                                v2.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(iIndex2)
                                dT = t2.pins(strDCVI_PinArray(intJ)).value - t1.pins(strDCVI_PinArray(intJ)).value
                                If dT = 0 Then
                                        t2.pins(strDCVI_PinArray(intJ)).value = t2.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        deltaV.AddPin (strTestDCVI_Pin)
        deltaV = v2.Math.Subtract(v1)
        deltaT.AddPin (strTestDCVI_Pin)
        deltaT = t2.Math.Subtract(t1)

        pldMeasureValue.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValue = deltaT.Math.Multiply(dblTestMeasureCurrent)
        pldMeasureValue = pldMeasureValue.Math.divide(deltaV)

        pldMeasureValue = pldMeasureValue.Math.Subtract(pldMeasureValueParasitic)

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmDefault

        Dim Tname As String
        '********************main test begin********************'
        'judgeCheckResult pldMeasureValue, 4714, "CV6|CV7|CV8, K96 ON"
       ' Tname = Test_Name & " Cap"
        Util_Judge_Result pldMeasureValue, tNum, "UVI80 cap", "F"
        'TheExec.Flow.TestLimit resultVal:=pldMeasureValue, PinName:=strTestDCVI_Pin, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone

        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************main test end********************'

        '********************subtest begin 1********************'
        'tNum=cstr(clng(tNum)+10)
        'judgeCheckResult pldMeasureValueParasitic, 4715, "K96 OFF" ', 0, 0.000000005, "F"
        Tname = Relay_On & " OFF Para Cap"
        Util_Judge_Result pldMeasureValueParasitic, tNum + 10, "UVI80 Parasitic cap", "F"
        'TheExec.Flow.TestLimit resultVal:=pldMeasureValueParasitic, PinName:=strTestDCVI_Pin, lowVal:=Para_Cap_Lo, hiVal:=Para_Cap_Hi, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone

        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************subtest end 1********************'

        Set pldMeasureValue = Nothing
        '********************capacitor test end********************'

        '***disconnect DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .Gate = False
                TheHdw.Wait 0.002
                .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 0.0005
        '***-----------------------***'
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function UVI80_Cap_Leakage(tNum As String, PowerPins As String)
On Error GoTo errHandler
    ' decompose power pins
    Dim PinArr() As String
    Dim PinCount As Long
    TheExec.DataManager.DecomposePinList PowerPins, PinArr(), PinCount
   
    Dim iMeasRange As Double: iMeasRange = 0.002  ' 2mA
    Dim sPin As Variant
    'TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
    For Each sPin In PinArr
        ' connect DCVI instrument
        TheHdw.DCVI.pins(sPin).Connect tlDCVIConnectDefault
        TheHdw.DCVI.pins(sPin).LocalKelvin = True
        TheHdw.Wait 0.0005

        ' set DCVI instrument
        With TheHdw.DCVI.pins(sPin)
            .mode = tlDCVIModeVoltage
            .ComplianceRange(tlDCVICompliancePositive).value = 7
            .ComplianceRange(tlDCVIComplianceNegative).value = 2
            TheHdw.Wait 0.01
            .Voltage = 0.1 ' ForceV_IiH
            .Current = iMeasRange
            .VoltageRange.value = 2
            .CurrentRange.value = iMeasRange
            .NominalBandwidth = 1003
            .Meter.mode = tlDCVIMeterCurrent
            .Meter.CurrentRange = iMeasRange
            .Meter.Filter.bypass = False
            .Meter.Filter.value = 10000
            .Alarm(tlDCVIAlarmMode) = tlAlarmOff
            .Gate = True
            TheHdw.Wait 0.01
            .SetCurrentAndRange 20 * uA, 20 * uA
            TheHdw.Wait 0.001
        End With

        ' main test
        Dim pld As New PinListData
        pld = TheHdw.DCVI.pins(sPin).Meter.Read(tlStrobe, 10, 100000, tlDCVIMeterReadingFormatAverage)
        Util_Judge_Result pld, tNum, "UVI80 Leakage", "A"

        ' disconnect DCVI instrument
        With TheHdw.DCVI.pins(sPin)
            .SetCurrentAndRange 0.01, 0.01
            .Gate = False
            TheHdw.Wait 0.002
            .Reset tlResetSettings + tlResetConnections
            TheHdw.Wait 0.0005
        End With
        
    Next sPin
 
    TheHdw.DCVI.pins(PowerPins).Alarm(tlDCVIAlarmMode) = tlAlarmDefault
    'TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
        
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in UVI80_Cap_Leakage"
End Function


Private Function UVI80_Leakage(tNum As String, PowerPins As String, Relay_On As String)
On Error GoTo errHandler
    ' decompose power pins
    Dim PinArr() As String
    Dim PinCount As Long
    TheExec.DataManager.DecomposePinList PowerPins, PinArr(), PinCount
   
    Dim iMeasRange As Double: iMeasRange = 0.002  ' 2mA
    Dim sPin As Variant
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
    For Each sPin In PinArr
        ' connect DCVI instrument
        TheHdw.DCVI.pins(sPin).Connect tlDCVIConnectDefault
        TheHdw.DCVI.pins(sPin).LocalKelvin = True
        TheHdw.Wait 0.0005

        ' set DCVI instrument
        With TheHdw.DCVI.pins(sPin)
            .mode = tlDCVIModeVoltage
            .ComplianceRange(tlDCVICompliancePositive).value = 7
            .ComplianceRange(tlDCVIComplianceNegative).value = 2
            TheHdw.Wait 0.01
            .Voltage = 0.1 ' ForceV_IiH
            .Current = iMeasRange
            .VoltageRange.value = 2
            .CurrentRange.value = iMeasRange
            .NominalBandwidth = 1003
            .Meter.mode = tlDCVIMeterCurrent
            .Meter.CurrentRange = iMeasRange
            .Meter.Filter.bypass = False
            .Meter.Filter.value = 10000
            .Alarm(tlDCVIAlarmMode) = tlAlarmOff
            .Gate = True
            TheHdw.Wait 0.01
            .SetCurrentAndRange 20 * uA, 20 * uA
            TheHdw.Wait 0.001
        End With

        ' main test
        Dim pld As New PinListData
        pld = TheHdw.DCVI.pins(sPin).Meter.Read(tlStrobe, 10, 100000, tlDCVIMeterReadingFormatAverage)
        Util_Judge_Result pld, tNum, "UVI80 Leakage", "A"

        ' disconnect DCVI instrument
        With TheHdw.DCVI.pins(sPin)
            .SetCurrentAndRange 0.01, 0.01
            .Gate = False
            TheHdw.Wait 0.002
            .Reset tlResetSettings + tlResetConnections
            TheHdw.Wait 0.0005
        End With
        
    Next sPin
 
    TheHdw.DCVI.pins(PowerPins).Alarm(tlDCVIAlarmMode) = tlAlarmDefault
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
    Exit Function

errHandler:
    TheExec.AddOutput "Error in the Seq Leakage Test"
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function DCVI_RLY_Rising_Meas(tNum As String, Test_PinName1 As String, Switch_Type As String, Relay_On As String)
On Error GoTo errHandler
    Dim Pins_On() As String
    Dim Pin_Cnt_On As Long
    Dim All_TestPin As String

    TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff

    ' connect DCVI instrument
    TheHdw.DCVI.pins(Test_PinName1).AlarmClear
    TheHdw.DCVI.pins(Test_PinName1).Connect tlDCVIConnectDefault
    TheHdw.DCVI.pins(Test_PinName1).Disconnect tlDCVIConnectHighSense
    TheHdw.DCVI.pins(Test_PinName1).LocalKelvin = True
    TheHdw.Wait 0.0005

    ' capacitor main test
    Dim j As Integer
    Dim deltaV As New PinListData
    Dim deltaT As New PinListData
    Dim dT As Double
    Dim DCVIResultDSP As New DSPWave
    Dim strDCVI_PinArray() As String
    Dim dblCapSampleRate As Double
    Dim dblCapSampleSize As Double
    Dim dblCurrentRange As Double
    Dim dblCaptureWait As Double
    Dim strTestDCVI_Pin As String
    Dim dblTestMeasureCurrent As Double
    Dim dblTestChargeVoltage As Double
    Dim dblTestOffsetVoltage As Double
    Dim iIndex1 As New SiteLong
    Dim iIndex2 As New SiteLong
    
    strTestDCVI_Pin = Test_PinName1
    dblTestMeasureCurrent = 0.0002
    strDCVI_PinArray = Split(strTestDCVI_Pin, ",")
    dblCapSampleRate = 1000 * kHz
    dblCapSampleSize = 512
    dblCaptureWait = dblCapSampleSize / dblCapSampleRate
    dblCurrentRange = 0.005

    TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmOff
    TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.Add ("CapSignal")
    With TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals("CapSignal")
        .SampleRate = 1000000
        .SampleSize = 256
        .LoadSettings
    End With

    ' set DCVI instrument
    With TheHdw.DCVI.pins(strTestDCVI_Pin)
        .mode = tlDCVIModeCurrent
        .ComplianceRange(tlDCVICompliancePositive).value = 7
        .ComplianceRange(tlDCVIComplianceNegative).value = 2
        TheHdw.Wait 0.01
        .Current = 0 '0.005
        .Voltage = 5
        .VoltageRange.value = 7
        .CurrentRange.Autorange = True
        .NominalBandwidth = 1003
        .Meter.mode = tlDCVIMeterVoltage
        .Meter.VoltageRange = 7
        .Meter.Filter.bypass = False
        .Meter.Filter.value = 10000
        .Alarm(tlDCVIAlarmMode) = tlAlarmOff
        .Gate = True
        TheHdw.Wait 0.01
    End With

    TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
    TheHdw.Wait 0.005

    ' measure parasitic capacitor begin
    Dim dblMeasureCurrentParasiticCap As Double
    Dim dblCurrentRangeParasiticCap As Double
    Dim pldV1_parasitic As New PinListData
    Dim pldV2_parasitic As New PinListData
    Dim pldDeltaV_parasitic As New PinListData
    Dim pldT1_parasitic As New PinListData
    Dim pldT2_parasitic As New PinListData
    Dim Swtich_DeltaT As New PinListData
    Dim pldDT_parasitic As Double
    Dim DCVIResultDSP_parasitic As New DSPWave
    Dim RLY_CapDSP As New DSPWave
    Dim i As Integer
    Dim Step2 As Integer
    Set Swtich_DeltaT = Nothing

    TheHdw.Wait 0.003
    Dim Start_time As Double
    Dim End_time As Double
    dblMeasureCurrentParasiticCap = 0.000001
    dblCurrentRangeParasiticCap = 0.000002

    'set the instrument to drive the constant current to charge the capacitor
    With TheHdw.DCVI.pins(strTestDCVI_Pin)
        .mode = tlDCVIModeCurrent
        .Current = dblMeasureCurrentParasiticCap
        .Voltage = 5
        TheHdw.Wait 0.01
    End With
    
    'start capture
    TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

    'switch on the gate and start source the constant current
    Start_time = TheExec.Timer(0)
    TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
    End_time = TheExec.Timer(Start_time)
    TheHdw.Utility.pins(Pins_On(0)).State = tlUtilBitOn

    TheHdw.Wait 0.00512
    TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
    TheHdw.Wait 0.005

    'computing of the parasitic cap value
    Dim vsite As Variant
    For j = 0 To UBound(strDCVI_PinArray)
        pldT1_parasitic.AddPin (strDCVI_PinArray(j))
        pldT2_parasitic.AddPin (strDCVI_PinArray(j))
        Swtich_DeltaT.AddPin (strDCVI_PinArray(j))

        For Each vsite In TheExec.sites.Active
            If TheExec.sites.Active Then
                DCVIResultDSP_parasitic = TheHdw.DCVI.pins(strDCVI_PinArray(j)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(j)).value
                iIndex1 = DCVIResultDSP_parasitic.FindIndex(OfFirstElement, GreaterThan, 0.01)
                Step2 = iIndex1 + 1
                For i = Step2 To 255
                    If DCVIResultDSP_parasitic.Element(i) < DCVIResultDSP_parasitic.Element(i - 1) Then
                        iIndex2 = i - 1
                        Exit For
                    End If
                Next i
                        
                pldT1_parasitic.pins(strDCVI_PinArray(j)).value = End_time
                pldT2_parasitic.pins(strDCVI_PinArray(j)).value = iIndex2 / dblCapSampleRate
                pldDT_parasitic = pldT2_parasitic.pins(strDCVI_PinArray(j)).value - pldT1_parasitic.pins(strDCVI_PinArray(j)).value
                Swtich_DeltaT = pldDT_parasitic
            End If
        Next vsite
    Next j

    ' judge result
    Util_Judge_Result Swtich_DeltaT, tNum, "UVI80 SW time", "S"

    ' clean up
    TheHdw.Utility.pins(Pins_On(0)).State = tlUtilBitOff
    With TheHdw.DCVI.pins(strTestDCVI_Pin)
        .mode = tlDCVIModeCurrent
        .Voltage = 0
        .Current = 0
        .Gate = True
        TheHdw.Wait 0.05
        .Gate = False
        TheHdw.Wait 0.005
    End With

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in DCVI_RLY_Rising_Meas"
End Function


'Ben Modify 20220307
'Cory Modify 20220307
Private Function DIBC_Lib_CheckUdbState(sRelayList As String, sUdbState As String) As PinListData
    On Error GoTo errHandler

    TheHdw.Utility.Threshold = 4
    
    'turn ON/ OFF relay and reab back its state
    If sUdbState = "ON" Then
        TheHdw.Utility.pins(sRelayList).State = tlUtilBitOn
    Else
        TheHdw.Utility.pins(sRelayList).State = tlUtilBitOff
    End If
    TheHdw.Wait 0.03
    
    Set DIBC_Lib_CheckUdbState = TheHdw.Utility.pins(sRelayList).States(tlUBStateCompared)

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in DIBC_Lib_CheckUdbState"
End Function


Private Function Relay_Circuit_check(tNum As String, Force_volt As Double, Current_Range As Double, Off_V_LowLimit As Double, Off_V_HiLimit As Double, Optional Relay_On As String, Optional WaitTime As Double = 0.003, Optional Route_Start As String, Optional Route_End As String)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "Relay_Circuit_check"

'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim Wait_Time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    'Dim R_Volt As PinListData
    Dim R_Value As New PinListData
    Dim F_volt As New PinListData
    Dim Res_Value As Double
    Dim FVolt As Double
    'Dim FCurr As Double
'''    Dim HiLimit As Double
'''    Dim LowLimit As Double
'    On Error GoTo errHandler
''''    HiLimit = 300#
''''    LowLimit = 200#
    'FVolt = 0.02
    'FCurr = 0.0004
    relayOnStr = ""
    relayOffStr = ""

    TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On


'''    Trim_NC_Pin Pins_On, Pin_Cnt_On
'''    Trim_NC_Pin Pins_Off, Pin_Cnt_Off



   'For Each site In TheExec.sites.Selected
     If Relay_On <> "" Then
        'theexec.Datalog.WriteComment "============================ Relay ON ====================================="
        Tname = "Site" & site & "-" & CStr(Relay_On) & " On"
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
        For Each p In Pins_On
                'Tname = "rly_on_" & p
                BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)

            'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
           ' TheExec.Datalog.WriteComment "Relay_on : " & BitState.Pins(p)
        Next p
     End If
          If Route_End <> "" Then
                If LCase(TheExec.DataManager.pinType(Route_End)) = "power" Then

                        If LCase(TheExec.DataManager.ChannelType(Route_End)) Like "*dcvs*" Then
                         TheHdw.DCVS.pins(Route_End).Connect tlDCVSConnectDefault
                         TheHdw.Wait 0.0005
                        '***-----------------------***'
                                    With TheHdw.DCVS.pins(Route_End)
                                       .mode = tlDCVSModeVoltage
                                       .Voltage.Main.value = 0
                                       .Voltage.Alt.value = 0
                                       .Voltage.Output = tlDCVSVoltageMain
                                       .CurrentRange.value = 0.002
                                       .Meter.mode = tlDCVSMeterCurrent
                                       .Meter.CurrentRange = 0.002
                                       .Alarm(tlDCVSAlarmAll) = tlAlarmOff
                                       .Gate = True
                                       TheHdw.Wait 0.25
                                    End With
                        ElseIf LCase(TheExec.DataManager.ChannelType(Route_End)) = "dcvi" Then
                              With TheHdw.DCVI.pins(Route_End)
                                     .mode = tlDCVIModeCurrent
                                     .ComplianceRange(tlDCVICompliancePositive).value = 7
                                     .ComplianceRange(tlDCVIComplianceNegative).value = 2
                                     TheHdw.Wait 0.01
                                     .Current = 0
                                     .Voltage = 0
                                     .CurrentRange = 0.02
                                     .VoltageRange = 2
                                     .NominalBandwidth = 1000
                                     .Meter.mode = tlDCVIMeterVoltage
                                     .Meter.Filter.bypass = False
                                     .Meter.Filter.value = 10000
                                     .Meter.VoltageRange = 2
                                     .Gate = True
                                     TheHdw.Wait 0.01
                             End With
                        End If
                ElseIf LCase(TheExec.DataManager.pinType(Route_End)) = "i/o" Then
                    TheHdw.Digital.pins(Route_End).Disconnect
                    With TheHdw.PPMU.pins(Route_End)
                        .Connect
                        .ForceV 0#
                        .Gate = tlOn

                    End With
                End If
                TheHdw.Wait 0.1
          Else


          End If


          If LCase(TheExec.DataManager.ChannelType(Route_Start)) = "i/o" Then
                    TheHdw.Digital.pins(Route_Start).Disconnect
                    With TheHdw.PPMU.pins(Route_Start)
                        .Connect
                        '.ForceI 0.002, 0.002
                        .ForceV Force_volt, Current_Range
                        .Gate = tlOn

                    End With

                     TheHdw.Wait 0.1
                   '----------------------------- Calculate Measure result

                    R_Value = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
                   ' R_Value = R_Value.Math.Divide(thehdw.PPMU.Pins(Route_Start).current.Value)
                    'F_volt.AddPin(Route_Start).Value = thehdw.Pins(Route_Start).PPMU.Voltage.Value
                    F_volt.AddPin (Route_Start)
                    F_volt.pins(Route_Start).value = TheHdw.pins(Route_Start).PPMU.Voltage.value
                    R_Value = F_volt.Math.divide(R_Value)
                    Util_Judge_Result R_Value, tNum, "Relay On Resistance", "Ohm"
                    tNum = CStr(CLng(tNum) + 1)
                     ' TheExec.Flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=Force_volt, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
           ElseIf LCase(TheExec.DataManager.ChannelType(Route_Start)) = "dcvi" Then
                        '***connect DCVI instrument***'

                With TheHdw.DCVI.pins(Route_Start)
                    .mode = tlDCVIModeCurrent
                    .Voltage = 6
                    .Current = 0.0005
                    .Connect tlDCVIConnectDefault
                    .Gate = True
                    .Meter.mode = tlDCVIMeterVoltage
                End With

                TheHdw.Wait 0.1

                '********************Resistor main test begin********************'
                'Dim pldMeasureValue As New PinListData

                R_Value = TheHdw.DCVI.pins(Route_Start).Meter.Read(tlStrobe, 2)
                R_Value = R_Value.Math.divide(0.0005)
                Util_Judge_Result R_Value, tNum, "Relay On Resistance", "Ohm"
                tNum = CStr(CLng(tNum) + 1)
                'TheExec.Flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=Force_volt, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone

                'Reset DCVI
                With TheHdw.DCVI.pins(Route_Start)
                    .SetCurrentAndRange 0.2, 0.2
                    .Current = 0
                    .CurrentRange.Autorange = True
                    .Voltage = 0
                    .Gate(tlDCVIGateHiZ) = False
                    .BleederResistor = tlDCVIBleederResistorAuto
                    .Disconnect
                    .mode = tlDCVIModeCurrent
                End With
           End If


            Set R_Value = Nothing
            Set F_volt = Nothing
           '----------------------------------------------------------------------------------------------------------------------------------------------------------

            If Relay_On <> "" Then



                Tname = "Site" & site & "-" & CStr(Relay_On) & " Off"
                TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
                '''TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)

                For Each p In Pins_On
                        'Tname = "rly_on_" & p
                        BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)

                        'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
                       ' TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
                Next p
                If LCase(TheExec.DataManager.ChannelType(Route_Start)) = "i/o" Then
                             TheHdw.PPMU.pins(Route_Start).ClampVHi = 1#

                             TheHdw.PPMU.pins(Route_Start).ForceV 0#

                             TheHdw.PPMU.pins(Route_Start).ForceV Force_volt, Current_Range

                             TheHdw.Wait 0.002
                             R_Value = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)

                             Dim counter As Integer: counter = 0
                             F_volt.AddPin (Route_Start)
                             For Each site In TheExec.sites.Selected
                                While R_Value.Math.compare(EqualTo, 0) And counter <= 10
                                  R_Value = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
                                  counter = counter + 1
                                Wend


                                F_volt.pins(Route_Start).value = TheHdw.pins(Route_Start).PPMU.Voltage.value
                                R_Value = F_volt.Math.divide(Abs(R_Value))
                             Next site
                             Util_Judge_Result R_Value, tNum, "Relay Off Resistance", "Ohm"
                            ' TheExec.Flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=Force_volt, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
                             Set R_Value = Nothing
                             Set F_volt = Nothing
                             With TheHdw.PPMU.pins(Route_Start)
                                 .ForceV (0)
                                 .Disconnect
                                 .Gate = tlOff

                             End With
                  ElseIf LCase(TheExec.DataManager.ChannelType(Route_Start)) = "dcvi" Then
                                             '***set DCVI instrument***'
                            With TheHdw.DCVI.pins(Route_Start)
                                    .Current = 0.0015
                                    .Voltage = 1
                                    .CurrentRange = 0.002
                                    .VoltageRange = 1
                                    .Meter.VoltageRange = 2
                                    .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                                    .Gate = False
                                    TheHdw.Wait 0.005
                            End With
                            '***-----------------------***'

                            TheHdw.Wait 0.003

                            TheHdw.DCVI.pins(Route_Start).Gate = True
                            TheHdw.Wait 0.003
                            '********************subtest begin 1********************'
                            R_Value = TheHdw.DCVI.pins(Route_Start).Meter.Read(tlStrobe, 100, 100000)
                            Util_Judge_Result R_Value, tNum, "Relay Off Resistance", "Ohm"
                         '  TheExec.Flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, hiVal:=Off_V_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=Force_volt, ForceUnit:=unitAmp, customUnit:="V", ForceResults:=tlForceNone

                  End If
                TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff

            End If
            If Route_End <> "" Then
                If LCase(TheExec.DataManager.pinType(Route_End)) = "power" Then
                     With TheHdw.DCVS.pins(Route_End)
                         .mode = tlDCVSModeVoltage
                         .Voltage.Main = 0#
                         .Disconnect
                         .Gate = False
                     End With
                 ElseIf LCase(TheExec.DataManager.pinType(Route_End)) = "i/o" Then
                     TheHdw.Digital.pins(Route_End).Disconnect
                     With TheHdw.PPMU.pins(Route_End)
                         .ForceV 0#
                         .Disconnect
                         .Gate = tlOff

                     End With
                End If
             Else

             End If


'   Next site
    TheHdw.Digital.pins("All_Digital").Disconnect
    TheHdw.PPMU.pins("All_Digital").Disconnect
'Exit Function



Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Sub Relay_Parasitic_Cap(pat As String, tNum As String, Tname As String, pin As String, relayOn As String, relayOff As String)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "Relay_Parasitic_Cap"

    Dim t       As Double: t = 20 * ns
    Dim iMeas   As Double: iMeas = 0.00001
    Dim vCharge As Double: vCharge = 0.8
    Dim vOffset As Double: vOffset = 0.3

    ' set relay
    If relayOn <> "" Then TheHdw.Utility.pins(relayOn).State = tlUtilBitOn
    If relayOff <> "" Then TheHdw.Utility.pins(relayOff).State = tlUtilBitOff

    ' ensure pin voltage equals to 0V
    TheHdw.Digital.pins(pin).Disconnect
    TheHdw.PPMU.pins(pin).Connect
    With TheHdw.PPMU.pins(pin)
        .ForceV 0
        .Gate = tlOn
        TheHdw.Wait 0.005
        .Gate = tlOff
        TheHdw.Wait 0.002
        .Disconnect
    End With
    TheHdw.Digital.ConnectPins (pin)
    TheHdw.Wait 0.005

    ' run pattern
    Dim pldFailCount As New PinListData
    Dim pldFailCount_SiteValue As New SiteDouble
    Dim pldFailCountTemp_SiteValue As New SiteDouble
    Dim pldFailCountTemp As New PinListData
    Dim i As Integer
    For i = 1 To 10
        TheHdw.Digital.patterns.pat(pat).Run pin & "_CX_100usPer020ns"  ' start label
           TheHdw.Wait 0.005
           If (i = 1) Then
               pldFailCount_SiteValue = TheHdw.Digital.pins(pin).FailCount
               pldFailCount.AddPin pin
               pldFailCount.pins(pin).value = pldFailCount_SiteValue
               pldFailCountTemp.AddPin pin ' add temp pin  here 20201230??
           Else
                pldFailCountTemp_SiteValue = TheHdw.Digital.pins(pin).FailCount


                pldFailCountTemp.pins(pin).value = pldFailCountTemp_SiteValue
                pldFailCount = pldFailCount.Math.Add(pldFailCountTemp)
           End If
           TheHdw.Wait 0.005

    Next i
    pldFailCount = pldFailCount.Math.divide(10)

    ' judge result
    'Util_Judge_Result failCount.Multiply(iMeas).Multiply(t).Divide(vCharge), tNum, tName, "F" 'mask it,change to below
    Util_Judge_Result pldFailCount.Math.Multiply(t).Multiply(iMeas).divide(vCharge), tNum, Tname, "F" 'new add 20210111

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Sub

Private Sub Relay_Parasitic_Caps(tNum As String, pin As String, k1 As String, k2 As String, k3 As String, k4 As String)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "Relay_Parasitic_Caps"

    ' load pattern
    Dim pat As String: pat = ".\\" & DIBC_FOLDER & "\\Patterns\\CX_HSD_" & pin & ".pat"
    Dim vOffset As Double: vOffset = 0.2
    Dim vCharge As Double: vCharge = 0.8
    Dim iMeas As Double: iMeas = 0.00001
    ' ignore pat already loaded error
    On Error Resume Next
    TheHdw.Digital.patterns.pat(pat).Load
'    On Error GoTo 0

    ' load level timing
    TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, RelayMode:=tlPowered, LoadLevels:=True, LoadTiming:=True, _
                                     PinLevelsSheet:="DIBC_Levels__FT", TimeSetSheet:="DIBC_TSets__FT"

    With TheHdw.PinLevels.pins(pin)
        .ModifyLevel chVih, 2
        .ModifyLevel chVil, -0.2
        .ModifyLevel chVol, vOffset
        .ModifyLevel chVch, 5.5
        .ModifyLevel chVcl, -1
        .ModifyLevel chVoh, vOffset + vCharge
        .ModifyLevel chIoh, 0
        .ModifyLevel chIol, iMeas
        .ModifyLevel chVt, 5.5
    End With

    ' meas cap
    Call Relay_Parasitic_Cap(pat, Util_Incr(tNum, 0), "DIFF_CAP", pin, relayOn:=Join(Array(k3, k4), ","), relayOff:=Join(Array(k1, k2), ","))

    Call Relay_Parasitic_Cap(pat, Util_Incr(tNum, 1), k3 & "+" & k4 & "_CAP", pin, relayOn:=k4, relayOff:=Join(Array(k1, k2, k3), ","))

    ' clean up
    With TheHdw.PPMU.pins(pin)  ' debug here
        .ForceV 0
        .Gate = tlOff
        TheHdw.Wait 0.002
        .Reset tlResetSettings
        .Disconnect
    End With

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Sub

Private Function DIBC_Lib_HEXVS_C_Open_MeasureCapacitor_2(tNum As String, sPinName As String, _
                                      sInstrumentType As String, _
                                      dbExpectedCapacitance As Double, _
                                      dbForceCurrent As Double, _
                                      dbForceVoltage As Double, _
                                      Optional dblParallelResistor As Double = -1, _
                                      Optional strParasiticRelayPin As String = vbNullString, _
                                      Optional blRelayON As Boolean = True)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DIBC_Lib_HEXVS_C_Open_MeasureCapacitor_2"

'   On Error GoTo errHandler

    Dim dbCaptureTime As Double: dbCaptureTime = 0.05 ' 50ms
    Dim dbIRange As Double
    Dim dSampleRate As Double
    Dim lCaptrueDepth As Long
    Dim dbMinForceCurrent As Double: dbMinForceCurrent = 0.00001 ' 10uA minimum current
    Dim dbMaxForceCurrent As Double: dbMaxForceCurrent = 0.002    ' 2mA max current
    Dim ilimit_onrange As Double: ilimit_onrange = 0
    Dim sdbCapValue As New SiteDouble
    Dim sdbCapValue_Parasitic As SiteDouble
    Dim pldMeasureValue As New PinListData

    If sInstrumentType = "HEXVS" Then
        dSampleRate = 25000000#
        lCaptrueDepth = 256000#
        dbIRange = 1
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.5
    ElseIf sInstrumentType = "HDVS" Then
        dSampleRate = 5555555#
        lCaptrueDepth = 2048#
        dbIRange = 1
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.1
    ElseIf sInstrumentType = "VHDVS" Then
        If dbExpectedCapacitance > 0.0001 Then
            dSampleRate = 100000#
        Else
            dSampleRate = 100000#
        End If
        lCaptrueDepth = 16384#
        dbIRange = 0.02
        ilimit_onrange = 0.125
    ElseIf sInstrumentType = "vs-800ma" Then
        If dbExpectedCapacitance > 0.0001 Then
            dSampleRate = 100000#
        Else
            dSampleRate = 100000#
        End If
        lCaptrueDepth = 16384#
        dbIRange = 0.02
        ilimit_onrange = 0.125
    ElseIf sInstrumentType = "VSM" Then
        dSampleRate = 30000000#
        lCaptrueDepth = 256000#
        dbIRange = 1
    ElseIf sInstrumentType = "vs-5a" Then
        dSampleRate = 100000#
        lCaptrueDepth = 16384#
        dbIRange = 0.2
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.5
    Else
        MsgBox "Unsupported DCVS type found : " & sInstrumentType
    End If

    ' If current is outside limits, set to limit and recalculate deltaT
    If dbForceCurrent < dbMinForceCurrent Then
        dbForceCurrent = dbMinForceCurrent
        dbCaptureTime = dbExpectedCapacitance * dbForceVoltage / dbForceCurrent
    ElseIf dbForceCurrent > dbMaxForceCurrent Then
        dbForceCurrent = dbMaxForceCurrent
        dbCaptureTime = dbExpectedCapacitance * dbForceVoltage / dbForceCurrent
    End If

    Dim lSampleNum As Long
    Dim sWaveSignalName As String: sWaveSignalName = "DIBC_DCVS_CapMeasureSig"
    Dim dspWave_VCapture As New DSPWave
    Dim dspWave_ICapture As New DSPWave
    Dim vLow As Double
    Dim vHigh As Double
    Dim lIndexlow As Long
    Dim lIndexHigh As Long
    Dim dbTimeDelta As Double
    Dim vsite As Variant

    vLow = 0.4 * dbForceVoltage
    vHigh = 0.7 * dbForceVoltage
    lSampleNum = CLng(dSampleRate * dbCaptureTime)
    If lSampleNum > lCaptrueDepth Then lSampleNum = lCaptrueDepth

    Dim savedAlarmSetting As tlAlarmBehavior
    Dim dbNewIRange As Double
    Dim rangeOptions() As Double
    Dim i As Long
    Dim rangeIncreasing As Boolean
    If TheExec.sites.Selected.Count = 0 Then Exit Function

    ' Connect the DCVS pin forcing 0V and source current clamps at dbForceCurrent
    With TheHdw.DCVS.pins(sPinName)
        .Alarm(tlDCVSAlarmAll) = tlAlarmOff ' disable alarms
        .Voltage.Main.value = 0
        dbNewIRange = dbIRange
        ' Get list of current range options
        rangeOptions = TheHdw.DCVS.pins(sPinName).CurrentRange.list

        If UBound(rangeOptions) >= 1 Then
            ' Are the values in the list increasing?
            If rangeOptions(0) < rangeOptions(1) Then
                rangeIncreasing = True
            Else
                rangeIncreasing = False
            End If

            ' Determine the new Range
            For i = 0 To UBound(rangeOptions)
                If rangeIncreasing Then 'range values are increasing
                    If dbForceCurrent <= rangeOptions(i) Then
                            dbNewIRange = rangeOptions(i)
                        Exit For
                    End If
                Else ' range values are decreasing
                    If dbForceCurrent > rangeOptions(i) Then
                        If i > 0 Then
                            dbNewIRange = rangeOptions(i - 1)
                        Else
                            dbNewIRange = rangeOptions(0)
                        End If
                        Exit For
                    End If
                End If
            Next i
        End If

        ' Are we increasing current?
        If TheHdw.DCVS.pins(sPinName).CurrentLimit.Source.FoldLimit.level.value < dbForceCurrent Then
            ' First increase the range, then the current
            TheHdw.DCVS.pins(sPinName).CurrentRange = dbNewIRange
            If sInstrumentType = "vs-5a" Then
            Else
                TheHdw.DCVS.pins(sPinName).CurrentRange = dbNewIRange
            End If
            TheHdw.DCVS.pins(sPinName).CurrentLimit.Source.FoldLimit.level.value = dbForceCurrent
        Else  ' we are decreasing current

            TheHdw.DCVS.pins(sPinName).CurrentRange = dbNewIRange
            TheHdw.DCVS.pins(sPinName).CurrentLimit.Source.FoldLimit.level.value = dbForceCurrent
            If sInstrumentType = "vs-800ma" Then
                TheHdw.DCVS.pins(sPinName).CurrentRange = dbNewIRange
            Else
                TheHdw.DCVS.pins(sPinName).CurrentRange = dbNewIRange

            End If


        End If
        .mode = tlDCVSModeVoltage
        .Meter.mode = tlDCVSMeterVoltage
        If (sInstrumentType = "vs-800ma") Then 'Or (sInstrumentType = "vs-5a") Then
            .CurrentRange = dbIRange
        ElseIf (sInstrumentType = "vs-5a") Then
            .CurrentRange = dbNewIRange
        Else
            .Meter.CurrentRange = dbIRange
        End If
        .Connect (tlDCVSConnectDefault)
        .Gate = True
        Call TheHdw.Wait(0.05)
    End With

    ' Set up the capture signal
    With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterVoltage
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lSampleNum
        .Signals(sWaveSignalName).LoadSettings
    End With

    Call TheHdw.Wait(0.01)

    ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger

    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage

    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend

    ' Read the capture wave and process the results
    dspWave_VCapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave

    ' Go back to 0V
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = 0
    TheHdw.Wait (2 * dbCaptureTime)
    'TheHdw.Wait (0.15)

    ' Set up current capture signal
    With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterCurrent
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lSampleNum
        .Signals(sWaveSignalName).LoadSettings
    End With

    Call TheHdw.Wait(0.001)

     ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger

    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage

    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend

    ' Read the capture wave and process the results
    dspWave_ICapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave

    For Each vsite In TheExec.sites.Selected
        lIndexlow = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vLow)
        lIndexHigh = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vHigh)
        sdbCapValue.value = -1
        If lIndexlow > 1 And lIndexHigh < dspWave_VCapture.SampleSize - 1 And lIndexlow < lIndexHigh Then
            dbTimeDelta = (lIndexHigh - lIndexlow) / dSampleRate
            dbForceCurrent = dspWave_ICapture.Select(lIndexlow, 1, lIndexHigh - lIndexlow).CalcMean
            If dblParallelResistor < 0 Then  'no parallel resistor
                'computing of the capacitance value: C= I * dt / dU
                sdbCapValue.value = dbForceCurrent * (dbTimeDelta / (vHigh - vLow))
            Else
                'computing of the capacitance value: C= dt * (I/dU - 1/R)
                sdbCapValue.value = dbTimeDelta * (dbForceCurrent / (vHigh - vLow) - 1 / dblParallelResistor)
            End If
        End If
    Next vsite

    Dim Tname As String
    Tname = "Power_Cap" '& sPinName
    pldMeasureValue.AddPin sPinName
    pldMeasureValue.pins(sPinName) = sdbCapValue
    Util_Judge_Result pldMeasureValue, tNum, Tname, "F"
   ' TheExec.Flow.TestLimit resultVal:=sdbCapValue, PinName:=sPinName, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=dbForceVoltage, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone

    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = 0
    TheHdw.Wait (0.1)
    TheHdw.DCVS.pins(sPinName).Disconnect
    TheHdw.DCVS.pins(sPinName).Gate = False

    TheHdw.Wait (0.15)
'    Exit Function


Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function UVS256_HexVs_Leakage(tNum As String, ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, Power_pins As String, Power_Type As String)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "UVS256_HexVs_Leakage"
    Dim PinArr() As String, PinCount As Long
    Dim sAllDCVSPins As String
    Dim strAllDCVSPins()  As String
    Dim strAllDCVSMerged2Pins()  As String
    Dim strAllDCVSMerged4Pins()  As String
    Dim strAllDCVSMerged8Pins()  As String
    Dim lngPnum As Long
    Dim Hexvs_Pins As String
    Dim i As Integer
    Dim sPin As Variant
    Dim lIndex As Long
    Dim TestName As String
    Dim Min_Current_Range As Double
    Dim Max_Current_Range As Double
    Dim pldMeasureValue As New PinListData
    Dim MeasValue_Temp As New SiteDouble
    Dim sWaveSignalName As String: sWaveSignalName = "DIBC_DCVS_CapMeasureSig"
    Dim dspWave_ICapture As New DSPWave

    lIndex = 0
    TheExec.DataManager.DecomposePinList Power_pins, PinArr(), PinCount
    If UCase(Power_Type) Like UCase("*HexVs*") Or UCase(Power_Type) Like UCase("*vs-800ma*") Or UCase(Power_Type) Like UCase("*vs-5a*") Then
             TestName = "HexVs Leakage"
             For Each sPin In PinArr
                    TheHdw.DCVS.pins(sPin).Connect tlDCVSConnectDefault

                     TheHdw.Wait 0.01
                     With TheHdw.DCVS.pins(sPin)
                        .Gate = True
                        .mode = tlDCVSModeVoltage
                        .Voltage.Main.value = ForceV_IiH
                        .Voltage.Alt.value = ForceV_IiH
                        .Voltage.Output = tlDCVSVoltageMain
                        TheHdw.Wait 0.5 '1#
                        .CurrentRange.value = I_Meas_Range
                    Min_Current_Range = .Meter.CurrentRange.Min
                        .Meter.mode = tlDCVSMeterCurrent
                        If Power_Type = "vs-800ma" Or UCase(Power_Type) Like UCase("*vs-5a*") Then
                            .CurrentRange = Min_Current_Range
                        Else
                            .Meter.CurrentRange = Min_Current_Range
                        End If
                        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
                        '.Gate = True
                        TheHdw.Wait 1 '1#

                        If Power_Type = "vs-800ma" Or UCase(Power_Type) Like UCase("*vs-5a*") Then
                            If UCase(sPin) Like UCase("*VDD_FABRIC*") Then
                                .CurrentRange = Min_Current_Range
                            Else
                            .CurrentRange = Min_Current_Range
                            'TheHdw.Wait 0.01
                            End If
                        End If
                    End With
                     '***-----------------------***'
                      TestName = sPin & " pin Leakage"
                     TheHdw.Wait 0.4
                     '********************main test begin********************'
                     MeasValue_Temp = TheHdw.DCVS.pins(sPin).Meter.Read(tlStrobe, 1000, 100000, tlDCVIMeterReadingFormatAverage)
                     pldMeasureValue.AddPin (sPin)
                     pldMeasureValue.pins(sPin) = MeasValue_Temp
                     '********************main test end********************'

                     '********************Power Supply leakage test end********************'
                    'TheExec.Flow.TestLimit resultVal:=pldMeasureValue, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TestName, ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                        Util_Judge_Result pldMeasureValue, tNum, TestName, "A"
''                    judgeCheckResult pldMeasureValue, 13600 + lIndex, "DCVS_Short"
''                    If Theexec.sites.ActiveCount = 0 Then Exit Sub
''
''                    lIndex = lIndex + 1

            Next sPin
    ElseIf UCase(Power_Type) Like UCase("*VS256*") Then

            TestName = "UVS256 Leakage"
            sPin = Power_pins  'one power pins use 20201230
            TheHdw.DCVS.pins(sPin).Connect tlDCVSConnectDefault
            TheHdw.Wait 0.1




                 Max_Current_Range = TheHdw.DCVS.pins(sPin).CurrentRange.Max
                If UCase(sPin) = "SPI_PWR" Or UCase(sPin) = "BUFFER_PWR_2_DOMAIN" Or UCase(sPin) = "VDD_AMPH" Or UCase(sPin) = "BUFFER_PWR" Then 'conenct to IC
                ElseIf UCase(sPin) = "VDD_GPU" Or UCase(sPin) = "VDD_SOC" Or UCase(sPin) = "VDD_PCPU" Or UCase(sPin) = "VDD1" Or UCase(sPin) = "VDD2" Or UCase(sPin) = "VDD_ECPU" Or UCase(sPin) = "VDD_CPU_SRAM" Then
                    'GoTo skip 'will execute later
                Else 'If UCase(sPin) = "VDD12_DN_PCIE" Then

'''''
                    TheHdw.DCVS.pins(sPin).Connect tlDCVSConnectDefault

                     TheHdw.Wait 0.1
                     With TheHdw.DCVS.pins(sPin)
                        .mode = tlDCVSModeVoltage

                        .Voltage.Main.value = ForceV_IiH
                        .Voltage.Alt.value = ForceV_IiH
                        .Voltage.Output = tlDCVSVoltageMain
                        .CurrentRange.value = 0.2 '.02
                        .Meter.mode = tlDCVSMeterCurrent
                        If Power_Type = "vs-800ma" Then
                            .CurrentRange = 0.2
                        Else
                        .Meter.CurrentRange = 0.2
                        End If
                        '.Meter.CurrentRange = 0.2
                        .Alarm(tlDCVSAlarmAll) = tlAlarmOff

                        .Gate = True


                       TheHdw.Wait 0.4
'
'
                        .CurrentRange.value = I_Meas_Range '0.000004
                        If Power_Type = "vs-800ma" Then
                            .CurrentRange = I_Meas_Range
                        Else
                            .Meter.CurrentRange = I_Meas_Range
                        End If
                        '.Meter.CurrentRange = I_Meas_Range '0.000004

                        TheHdw.Wait 0.4
                      End With
                     '***-----------------------***'
                      TestName = " pin Leakage" ' sPin & " pin Leakage"
                    'Dim pldMeasureValue As New PinListData

                     '********************main test begin********************'
'''''                       dspWave_ICapture = thehdw.DCVS.Pins(sPin).Capture.Signals(sWaveSignalName).DSPWave
                     MeasValue_Temp = TheHdw.DCVS.pins(sPin).Meter.Read(tlStrobe, 1000, 100000, tlDCVIMeterReadingFormatAverage)
                     pldMeasureValue.AddPin (sPin)
                     pldMeasureValue.pins(sPin) = MeasValue_Temp

                     Util_Judge_Result pldMeasureValue, tNum, TestName, "A"
                End If

    End If
    Set pldMeasureValue = Nothing

    ' disconnect DCVS instrument
    With TheHdw.DCVS.pins(Power_pins)
       .CurrentRange.value = 0.02
        If (Power_Type = "vs-800ma") Or (Power_Type = "vs-5a") Then
            .CurrentRange = 0.02
        Else
       .Meter.CurrentRange = 0.02
        End If
       '.Meter.CurrentRange = 0.02
       .Voltage.Main = 0#
       .Voltage.Alt = 0#
       .Gate = False
       TheHdw.Wait 0.002
       .Disconnect tlDCVSConnectDefault
       TheHdw.Wait 0.0005
    End With

    Set pldMeasureValue = Nothing
'    Exit Function


Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function DIBC_Lib_UP1600_IO_Open(tNum As String, iopins As Variant)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DIBC_Lib_UP1600_IO_Open"
'    On Error GoTo errHandler

    Dim site     As Variant
    Dim PinArr() As String
    Dim PinCount As Long
    Dim i As Long
    Dim p As Variant
    Dim MeasVal As New PinListData
    Dim TestNum As Long
    Dim Tname As String
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim ForceV_IiH As Double: ForceV_IiH = 5#
    Dim ForceV_IiL As Double: ForceV_IiL = 5#
    Dim I_Meas_Range As Double: I_Meas_Range = 0.000002
    Dim leakage_pins As String: leakage_pins = Join(iopins, ",")
    Dim Relay_On As String: Relay_On = ""

    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
    TheHdw.PPMU.pins(leakage_pins).ForceV (0)
    TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
    TheHdw.PPMU.pins(leakage_pins).Disconnect

    ''''''use the "theexec.DataManager.DecomposePinList" to serialize the pins to be tested sequentially'''''
    TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
'    DCVS_Trim_NC_Pin PinArr(), PinCount

    ''High

    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    TheHdw.Digital.pins(leakage_pins).Disconnect

    With TheHdw.PPMU(leakage_pins)

        .Connect
        .Gate = tlOn
        .ForceV ForceV_IiH, I_Meas_Range
         TheHdw.Wait 0.01
         'DebugPrintFunc_PPMU leakage_pins.Value
         MeasVal = .Read(tlPPMUReadMeasurements)

        .ForceV (0)
        .Gate = tlOff
        .Disconnect
    End With

    'TheHdw.digital.Pins(leakage_pins).Connect 'connect the tested pin back to the PE

    'offline mode simulation
    If TheExec.TesterMode = testModeOffline Then

        For Each site In TheExec.sites
            For Each p In PinArr()
                If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.pins(p).value = 5 * uA + Rnd() * 0.1 * uA
            Next p
        Next site
    End If
    Util_Judge_Result MeasVal, tNum, "UP1600_IO_Leakage", "A"
'Set PC_Leakage_new = MeasVal

   ' TheExec.Flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:="Leakage", forceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone

TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff

'Exit Function

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function DIBC_Lib_UP2200_IO_Open(tNum As String, iopins As Variant)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DIBC_Lib_UP2200_IO_Open"

'    On Error GoTo errHandler

    Dim site     As Variant
    Dim PinArr() As String
    Dim PinCount As Long
    Dim i As Long
    Dim p As Variant
    Dim MeasVal As New PinListData
    Dim TestNum As Long
    Dim Tname As String
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim ForceV_IiH As Double: ForceV_IiH = 5#
    Dim ForceV_IiL As Double: ForceV_IiL = 5#
    Dim I_Meas_Range As Double: I_Meas_Range = 0.000002
    Dim leakage_pins As String: leakage_pins = Join(iopins, ",")
    Dim Relay_On As String: Relay_On = ""

    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
    TheHdw.PPMU.pins(leakage_pins).ForceV (0)
    TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
    TheHdw.PPMU.pins(leakage_pins).Disconnect

    ''''''use the "theexec.DataManager.DecomposePinList" to serialize the pins to be tested sequentially'''''
    TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
'    DCVS_Trim_NC_Pin PinArr(), PinCount

    ''High

    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    TheHdw.Digital.pins(leakage_pins).Disconnect

    With TheHdw.PPMU(leakage_pins)

        .Connect
        .Gate = tlOn
        .ForceV ForceV_IiH, I_Meas_Range
         TheHdw.Wait 0.01
         'DebugPrintFunc_PPMU leakage_pins.Value
         MeasVal = .Read(tlPPMUReadMeasurements)

        .ForceV (0)
        .Gate = tlOff
        .Disconnect
    End With

    'TheHdw.digital.Pins(leakage_pins).Connect 'connect the tested pin back to the PE

    'offline mode simulation
    If TheExec.TesterMode = testModeOffline Then

        For Each site In TheExec.sites
            For Each p In PinArr()
                If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.pins(p).value = 5 * uA + Rnd() * 0.1 * uA
            Next p
        Next site
    End If
    Util_Judge_Result MeasVal, tNum, "UP1600_IO_Leakage", "A"
'Set PC_Leakage_new = MeasVal

   ' TheExec.Flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:="Leakage", forceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone

TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff

'Exit Function

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function Digital_Open_Voltage_Meas(tNum As String, iopins As Variant, RelayOn_Pin As String) As Long
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "Digital_Open_Voltage_Meas"

Dim MeasureValue As New PinListData
Dim pldMeasureValue As New PinListData

Dim pldOpenValue As New PinListData
Dim ForceValI As Double
Dim MeasurePinStr As String

Dim PinArr() As String
Dim PinCount As Long
Dim pin  As Variant
Dim Res_Value As Double
Dim i As Integer
'Dim HiLimit As Double
'Dim LowLimit As Double
Dim force_i As Double
Dim Tname As String
Dim Relay_on_pin_temp1() As String 'add 191121
Dim Relay_on_pin_temp2() As String 'add 191121
Dim Relay_Name As String
Dim IOpins_temp As String: IOpins_temp = Join(iopins, ",")
'On Error GoTo errHandler
'HiLimit = 1.2
'LowLimit = 0.8
Dim sPin As String
'MeasurePinStr = TestPin
Tname = "Digital_Open"
'PinArr = iopins
TheExec.DataManager.DecomposePinList IOpins_temp, PinArr(), PinCount
Dim Current_Range As Double
Relay_on_pin_temp1 = Split(RelayOn_Pin, ";")  'add 191121
For i = 0 To UBound(Relay_on_pin_temp1)
    Relay_on_pin_temp2 = Split(Relay_on_pin_temp1(i), "+")
    Relay_On_Pin_Dic_open(Relay_on_pin_temp2(0)) = Relay_on_pin_temp2(1)
Next i
'TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
For i = 0 To UBound(PinArr)
sPin = PinArr(i)
'For Each pin In PinArr
 ' If UCase(PinArr(i)) = "REFCLK_XI0" Or UCase(PinArr(i)) = "REFCLK_RT_CLK32768" Or UCase(PinArr(i)) = "FRC_BACKUP_1" Or UCase(PinArr(i)) = "FRC_BACKUP_4" _
            Or UCase(PinArr(i)) = "FRC_BACKUP_5" Or UCase(PinArr(i)) = "FRC_BACKUP_2" Or UCase(PinArr(i)) = "REFCLK_MONITORCHANNEL" Or UCase(PinArr(i)) = "FRC_BACKUP_3" Then
            'these pin connected to clock buffer
  'ElseIf UCase(PinArr(i)) = "VDD_PCPU_MONITOR" Or UCase(PinArr(i)) = "VDD_GPU_MONITOR" Or UCase(PinArr(i)) = "VDD_SOC_MONITOR" Or UCase(PinArr(i)) = "VDD_CPU_Monitor" Then
           'short with DCVI/DCVS pins, connect capacitors to gnd. Not test current leakage.
 ' Else 'Remove If 20200410
           MeasureValue.AddPin (PinArr(i))

          If Relay_On_Pin_Dic_open.Exists(PinArr(i)) Then
                Relay_Name = Relay_On_Pin_Dic_open(PinArr(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOn
                TheHdw.Wait 0.0001

        End If

      TheHdw.Digital.pins(PinArr(i)).Disconnect

'-------------------------------------------------
      With TheHdw.PPMU(PinArr(i))
           .Connect
          TheHdw.Wait 0.002
          '.SetClampsVHi (1)
          .ClampVHi = 1#
          .ForceI 0.0015, 0.0015
          .Gate = tlOn

          TheHdw.Wait 0.004

      '    .ForceV 0.1, 0.04

      End With
'-------------------------------------------------
          'If UCase(PinArr(i)) = "VDD_CPU_MONITOR" Then

          'Else

                pldMeasureValue = TheHdw.PPMU.pins(PinArr(i)).Read(tlPPMUReadMeasurements)
                pldOpenValue.AddPin (PinArr(i))
                pldOpenValue.pins(PinArr(i)) = pldMeasureValue.pins(PinArr(i))
          'End If
            With TheHdw.PPMU.pins(PinArr(i))

              .Gate = tlOff
              .Disconnect
            End With

          If Relay_On_Pin_Dic_open.Exists(PinArr(i)) Then
                Relay_Name = Relay_On_Pin_Dic_open(PinArr(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOff
                TheHdw.Wait 0.0001
          End If

    Next i

    Util_Judge_Result pldOpenValue, tNum, "UP1600_IO_Open_Volt", "V"
'    Exit Function


Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function DCVI_Cap_Meas_3(tNum As String, Test_PinName1 As String, Test_PinName2 As String, Optional Relay_On As String)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DCVI_Cap_Meas_3"

        Dim Pins_On() As String, Pin_Cnt_On As Long
        Dim All_TestPin As String
        'Call SmartRelaySwitch("")
        TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff

        All_TestPin = Test_PinName1 & "," & Test_PinName2

        '***connect DCVI instrument***'
        TheHdw.DCVI.pins(All_TestPin).AlarmClear
        TheHdw.DCVI.pins(All_TestPin).Connect tlDCVIConnectDefault
        TheHdw.DCVI.pins(All_TestPin).Disconnect tlDCVIConnectHighSense
        TheHdw.DCVI.pins(All_TestPin).LocalKelvin = True

        TheHdw.Wait 0.0005
        '***-----------------------***'


        '***set DCVI instrument***'
        With TheHdw.DCVI.pins(Test_PinName2)
                .mode = tlDCVIModeVoltage
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Voltage = 0
                .Current = 0.02
                .VoltageRange.value = 0.05
                .CurrentRange.value = 0.02
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 0
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmDefault
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'


        '********************capacitor main test begin********************'
        Dim vsite As Variant
        Dim intI As Integer
        Dim intJ As Integer
        Dim pldMeasureValue As New PinListData
        Dim v1 As New PinListData
        Dim v2 As New PinListData
        Dim deltaV As New PinListData
        Dim t1 As New PinListData
        Dim t2 As New PinListData
        Dim deltaT As New PinListData
        Dim dT As Double
        Dim DCVIResultDSP As New DSPWave
        Dim strDCVI_PinArray() As String
        Dim dblCapSampleRate As Double
        Dim dblCapSampleSize As Double
        Dim dblCurrentRange As Double
        Dim dblCaptureWait As Double
        Dim strTestDCVI_Pin As String
        Dim dblTestMeasureCurrent As Double
        Dim dblTestChargeVoltage As Double
        Dim dblTestOffsetVoltage As Double
        Dim iIndex1 As New SiteLong
        Dim iIndex2 As New SiteLong

        strTestDCVI_Pin = Test_PinName1
        dblTestMeasureCurrent = 0.0002
        dblTestChargeVoltage = 0.4
        dblTestOffsetVoltage = 0.2

        strDCVI_PinArray = Split(strTestDCVI_Pin, ",")

        dblCapSampleRate = 100 * kHz
        dblCapSampleSize = 512
        dblCaptureWait = dblCapSampleSize / dblCapSampleRate
        dblCurrentRange = 0.005

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmOff

        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.Add ("CapSignal")
        With TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals("CapSignal")
                .SampleRate = 100000
                .SampleSize = 512
                .LoadSettings
        End With

        '***set DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Current = 0.005
                .Voltage = 5
                .VoltageRange.value = 7
                .CurrentRange.Autorange = True
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 7
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'

        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '*****measure parasitic capacitor begin*****'

        Dim dblMeasureCurrentParasiticCap As Double
        Dim dblCurrentRangeParasiticCap As Double
        Dim pldMeasureValueParasitic As New PinListData
        Dim pldV1_parasitic As New PinListData
        Dim pldV2_parasitic As New PinListData
        Dim pldDeltaV_parasitic As New PinListData
        Dim pldT1_parasitic As New PinListData
        Dim pldT2_parasitic As New PinListData
        Dim pldDeltaT_parasitic As New PinListData
        Dim pldDT_parasitic As Double
        Dim DCVIResultDSP_parasitic As New DSPWave

        TheHdw.Utility.pins(Pins_On(0)).State = tlUtilBitOff
        TheHdw.Wait 0.003

        dblMeasureCurrentParasiticCap = 0.000002
        dblCurrentRangeParasiticCap = 0.000002

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblMeasureCurrentParasiticCap
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait 0.00512
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'computing of the parasitic cap value
        For intJ = 0 To UBound(strDCVI_PinArray)
                pldV1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldV2_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT2_parasitic.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
''''                                        Call DCVIResultDSP_parasitic.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP_parasitic = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

'        DCVIResultDSP_parasitic.Plot

                                iIndex1 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                pldV1_parasitic.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic.data(iIndex1)
                                pldV2_parasitic.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic.data(iIndex2)
                                pldDT_parasitic = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value - pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value
                                If pldDT_parasitic = 0 Then
                                        pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        pldDeltaV_parasitic.AddPin (strTestDCVI_Pin)
        pldDeltaV_parasitic = pldV2_parasitic.Math.Subtract(pldV1_parasitic)
        pldDeltaT_parasitic.AddPin (strTestDCVI_Pin)
        pldDeltaT_parasitic = pldT2_parasitic.Math.Subtract(pldT1_parasitic)

        pldMeasureValueParasitic.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValueParasitic = pldDeltaT_parasitic.Math.Multiply(dblMeasureCurrentParasiticCap)
        pldMeasureValueParasitic = pldMeasureValueParasitic.Math.divide(pldDeltaV_parasitic)

        For intJ = 0 To UBound(strDCVI_PinArray)
                If TheExec.TesterMode = testModeOffline Then
                        For Each vsite In TheExec.sites.Active
                                pldMeasureValueParasitic.pins(strDCVI_PinArray(intJ)).value = 0
                        Next vsite
                End If
        Next intJ

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.002

        TheHdw.Utility.pins(Pins_On(1)).State = tlUtilBitOff
        TheHdw.Wait 0.003

        '*****measure parasitic capacitor end*****'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '*****measure parasitic capacitor begin*****'

        Dim dblMeasureCurrentParasiticCap1 As Double
        Dim dblCurrentRangeParasiticCap1 As Double
        Dim pldMeasureValueParasitic1 As New PinListData
        Dim pldV1_parasitic1 As New PinListData
        Dim pldV2_parasitic1 As New PinListData
        Dim pldDeltaV_parasitic1 As New PinListData
        Dim pldT1_parasitic1 As New PinListData
        Dim pldT2_parasitic1 As New PinListData
        Dim pldDeltaT_parasitic1 As New PinListData
        Dim pldDT_parasitic1 As Double
        Dim DCVIResultDSP_parasitic1 As New DSPWave

        'strTestDCVI_Pin = Test_PinName2

        dblMeasureCurrentParasiticCap1 = 0.000002
        dblCurrentRangeParasiticCap1 = 0.000002

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblMeasureCurrentParasiticCap1
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait 0.00512
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'computing of the parasitic cap value
        For intJ = 0 To UBound(strDCVI_PinArray)
                pldV1_parasitic1.AddPin (strDCVI_PinArray(intJ))
                pldV2_parasitic1.AddPin (strDCVI_PinArray(intJ))
                pldT1_parasitic1.AddPin (strDCVI_PinArray(intJ))
                pldT2_parasitic1.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
'''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
'''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
'''''                                        Call DCVIResultDSP_parasitic1.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
'''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP_parasitic1 = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

'        DCVIResultDSP_parasitic1.Plot

                                iIndex1 = DCVIResultDSP_parasitic1.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP_parasitic1.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                pldT1_parasitic1.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                pldT2_parasitic1.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                pldV1_parasitic1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic1.data(iIndex1)
                                pldV2_parasitic1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic1.data(iIndex2)
                                pldDT_parasitic1 = pldT2_parasitic1.pins(strDCVI_PinArray(intJ)).value - pldT1_parasitic1.pins(strDCVI_PinArray(intJ)).value
                                If pldDT_parasitic1 = 0 Then
                                        pldT2_parasitic1.pins(strDCVI_PinArray(intJ)).value = pldT2_parasitic1.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        pldDeltaV_parasitic1.AddPin (strTestDCVI_Pin)
        pldDeltaV_parasitic1 = pldV2_parasitic1.Math.Subtract(pldV1_parasitic)
        pldDeltaT_parasitic1.AddPin (strTestDCVI_Pin)
        pldDeltaT_parasitic1 = pldT2_parasitic1.Math.Subtract(pldT1_parasitic)

        pldMeasureValueParasitic1.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValueParasitic1 = pldDeltaT_parasitic1.Math.Multiply(dblMeasureCurrentParasiticCap)
        pldMeasureValueParasitic1 = pldMeasureValueParasitic1.Math.divide(pldDeltaV_parasitic)

        For intJ = 0 To UBound(strDCVI_PinArray)
                If TheExec.TesterMode = testModeOffline Then
                        For Each vsite In TheExec.sites.Active
                                pldMeasureValueParasitic1.pins(strDCVI_PinArray(intJ)).value = 0
                        Next vsite
                End If
        Next intJ

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.02
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.002
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
        TheHdw.Wait 0.003

        '***connect DCVI instrument***'
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P,PAD_MTR_VREF_N").Disconnect tlDCVIConnectHighSense
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P,PAD_MTR_VREF_N").LocalKelvin = False
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P,PAD_MTR_VREF_N").Connect tlDCVIConnectDefault
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P").Disconnect tlDCVIConnectHighSense
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P").LocalKelvin = True
'        TheHdw.Wait 0.0005
        '***-----------------------***'


        '*****measure parasitic capacitor end*****'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '***set DCVI instrument***'
        With TheHdw.DCVI.pins(Test_PinName2)
                .mode = tlDCVIModeVoltage
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Voltage = 0
                .Current = 0.02
                .VoltageRange.value = 0.05
                .CurrentRange.value = 0.02
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 0
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmDefault
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'


        strTestDCVI_Pin = Test_PinName1



        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblTestMeasureCurrent
                .Voltage = 4
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait dblCaptureWait
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False

        'computing of the capacitor value
        For intJ = 0 To UBound(strDCVI_PinArray)
                v1.AddPin (strDCVI_PinArray(intJ))
                v2.AddPin (strDCVI_PinArray(intJ))
                t1.AddPin (strDCVI_PinArray(intJ))
                t2.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
''''                                        Call DCVIResultDSP.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

'        DCVIResultDSP.Plot

                                iIndex1 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                t1.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                t2.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                v1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(iIndex1)
                                v2.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(iIndex2)
                                dT = t2.pins(strDCVI_PinArray(intJ)).value - t1.pins(strDCVI_PinArray(intJ)).value
                                If dT = 0 Then
                                        t2.pins(strDCVI_PinArray(intJ)).value = t2.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        deltaV.AddPin (strTestDCVI_Pin)
        deltaV = v2.Math.Subtract(v1)
        deltaT.AddPin (strTestDCVI_Pin)
        deltaT = t2.Math.Subtract(t1)

        pldMeasureValue.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValue = deltaT.Math.Multiply(dblTestMeasureCurrent)
        pldMeasureValue = pldMeasureValue.Math.divide(deltaV)

        pldMeasureValue = pldMeasureValue.Math.Subtract(pldMeasureValueParasitic)

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmDefault
        Dim Tname As String

        '********************main test begin********************'
        'judgeCheckResult pldMeasureValue, 4716, "CV1|CV2, K97/K98 ON"
        ' Tname = Test_Name & " Cap"
         Util_Judge_Result pldMeasureValue, tNum, "UVI80 cap", "F"
        ' TheExec.Flow.TestLimit resultVal:=pldMeasureValue, PinName:=Test_PinName1, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone


        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************main test end********************'

        '********************subtest begin 1********************'
        'judgeCheckResult pldMeasureValueParasitic, 4717, "K97 OFF" ', 0, 0.000000005, "F"
        ' Tname = Pins_On(0) & " OFF Para Cap"
        tNum = CStr(CLng(tNum) + 10)
         Util_Judge_Result pldMeasureValueParasitic, tNum, "UVI80 Parasitic cap", "F"
        ' TheExec.Flow.TestLimit resultVal:=pldMeasureValueParasitic, PinName:=Test_PinName1, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone

        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************subtest end 1********************'

        '********************subtest begin 2********************'
       ' judgeCheckResult pldMeasureValueParasitic1, 4718, "K98 OFF"
       ' Tname = Pins_On(1) & " OFF Para Cap"
       tNum = CStr(CLng(tNum) + 20)
        Util_Judge_Result pldMeasureValueParasitic1, tNum, "UVI80 Parasitic cap1", "F"
        ' TheExec.Flow.TestLimit resultVal:=pldMeasureValueParasitic1, PinName:=Test_PinName1, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone

        If TheExec.sites.ActiveCount = 0 Then Exit Function

        '********************CAPACITOR subtest end 2********************'
        Set pldMeasureValue = Nothing
        '***disconnect DCVI instrument***'
        With TheHdw.DCVI.pins(All_TestPin)
                .Gate = False
                TheHdw.Wait 0.002
                .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 0.0005
        '***-----------------------***'
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function


'Ben Modify 20220307
'Cory Modify 20220307
Private Function ReadPinMapSheet() As Boolean
    Dim sheetnames() As String
    Dim pinmapname As String
    Dim Worksheet As Excel.Worksheet
    Dim PinName As String
    Dim pingroupname As String
    Dim i As Integer
    Set g_map_PinGroup = CreateObject("Scripting.Dictionary")
        
    sheetnames() = TheExec.job.GetSheetNamesOfType(DMGR_SHEEt_Type_PINMAP)
    pinmapname = sheetnames(0)
    Set Worksheet = ActiveWorkbook.Sheets(pinmapname)
    Worksheet.Activate
    
    For i = 4 To Worksheet.UsedRange.Rows.Count
        PinName = Worksheet.Cells(i, 3).value
        pingroupname = Worksheet.Cells(i, 2).value
        PinName = UCase(Trim(PinName))
        pingroupname = UCase(Trim(pingroupname))
        
        If g_map_PinGroup.Exists(pingroupname) = False Then
            g_map_PinGroup.Add pingroupname, CreateObject("Scripting.Dictionary")
        End If
        g_map_PinGroup(pingroupname).Add PinName, PinName
    Next i
End Function

'Ben Modify 20220307
'Cory Modify 20220307
Private Function ExistsInPinGroup(vPin As Variant, vPinGroup As String) As Integer
    ExistsInPinGroup = -1
    
    If g_ReadPinMap = False Then
        Call ReadPinMapSheet    'First time will read pinmap into memory first
        g_ReadPinMap = True
    End If
    
    If g_map_PinGroup.Exists(UCase(vPinGroup)) = False Then
        ExistsInPinGroup = -1
    ElseIf g_map_PinGroup(UCase(vPinGroup)).Exists(UCase(vPin)) Then
        ExistsInPinGroup = 1
    Else
        ExistsInPinGroup = 0
    End If
End Function

Private Sub Util_Meas_SB_Power(tNum As String, PowerValue As Variant, vForce As Double)
On Error GoTo errHandler
    Dim loLim As Double: loLim = vForce * 0.95
    Dim hiLim As Double: hiLim = vForce * 1.05
    Dim Component_Str As String
    ' meas power reading
    
    TheHdw.DIB.power.item(PowerValue).State = tlOn
    TheHdw.Wait 0.1
    Dim MeasV As Double: MeasV = TheHdw.DIB.power.item(PowerValue).Reading
    'thehdw.DIB.power.Item(powerValue).State = tlOff
    
    ' distribute to all site
    Dim meas As New PinListData: meas.AddPin CStr(PowerValue)
    Dim site As Variant
    For Each site In TheExec.sites.Active
       meas.pins(CStr(PowerValue)).value(site) = MeasV
    Next
    Component_Str = "SB_Power_" & PowerValue
    ' judge result
    Util_Judge_Result meas, tNum, Component_Str, "V"

Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "error in Util_Meas_SB_Power"
End Sub


Private Function DIBC_Lib_LOOP_BACK_IO_Leakage_Open(tNum As String, leakage_pins_ary As Variant)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DIBC_Lib_LOOP_BACK_IO_Leakage_Open"

'    On Error GoTo errHandler
    Dim site As Variant
    Dim PinArr() As String, PinCount As Long, i As Long
    Dim p As Variant
    Dim MeasVal As New PinListData
    Dim AllMeasureVal As New PinListData
    Dim TestNum As Long
    Dim Tname As String
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim Tname_str As String
    Dim ForceV_IiH As Double:     ForceV_IiH = 5#
    Dim ForceV_IiL As Double:     ForceV_IiL = 5#
    Dim I_Meas_Range As Double:     I_Meas_Range = 0.000002
    Dim Relay_On As String: Relay_On = ""
    Dim leakage_pins As String: leakage_pins = Join(leakage_pins_ary, ",")

    Tname_str = "Loopback Pin Leakage Meas"

    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
    TheHdw.PPMU.pins(leakage_pins).ForceV (0)
    TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
    TheHdw.PPMU.pins(leakage_pins).Disconnect

    ''''''use the "theexec.DataManager.DecomposePinList" to serialize the pins to be tested sequentially'''''
    TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
    TheHdw.Digital.pins(leakage_pins).Disconnect
    For Each p In PinArr
        With TheHdw.PPMU(p)
            .Connect
            .Gate = tlOn
            .ForceV ForceV_IiH, I_Meas_Range
            TheHdw.Wait 0.5
            MeasVal = .Read(tlPPMUReadMeasurements)
            .ForceV (0)
            .Gate = tlOff
            .Disconnect
            AllMeasureVal.AddPin (p)
            AllMeasureVal.pins(p) = MeasVal.pins(p)
        End With
    Next p

    Util_Judge_Result AllMeasureVal, tNum, "LOOP_BACK_pin_Leakage", "A"
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
'    Exit Function


Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function DIBC_Lib_LoopBack_1MHz_VOH_check(tNum As String, Loop_Back_pins_ary As Variant)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DIBC_Lib_LoopBack_1MHz_VOH_check"
    Dim pldMeasureValueVOH As New PinListData
    Dim strTestPattern As String
    Dim strPatternPath As String
    Dim patTestPattern As New Pattern
    If dibonTesterType = "Jaguar" Then
        strTestPattern = "MeasFreq.pat"
    ElseIf dibonTesterType = "UltraFLEXplus" Then

        strTestPattern = "MeasFreq.patx"
    End If

    strPatternPath = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\" & strTestPattern
    patTestPattern.value = strPatternPath
    Dim strTestPinlist As String
    Dim strSourePinlist As String
    Dim plstTestPinlist As New PinList

    Dim strTestPattern_waveform As String
    Dim strPatternPath_waveform As String
    Dim patTestPattern_waveform As New Pattern
    Dim Loop_Back_pins As String
    If dibonTesterType = "Jaguar" Then
        strTestPattern_waveform = "MeasFreq_waveform.pat"
    ElseIf dibonTesterType = "UltraFLEXplus" Then

        strTestPattern_waveform = "MeasFreq_waveform.patx"
    End If

    strPatternPath_waveform = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\" & strTestPattern_waveform
    patTestPattern_waveform.value = strPatternPath_waveform

    strTestPinlist = Loop_Back_pins_ary(1)
    strSourePinlist = Loop_Back_pins_ary(0)
    Loop_Back_pins = Join(Loop_Back_pins_ary, ",")
    TheHdw.Digital.pins(Loop_Back_pins).Connect
    TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, PinLevelsSheet:="DIBC_Levels__FT", TimeSetSheet:="DIBC_TSets__FT"  '0831 by rita 'DIBC_TSets__
    TheHdw.Wait 0.01

    With TheHdw.Digital.pins(Loop_Back_pins).Levels
        .value(chVil) = 0
        .value(chVih) = 0.9     '2
        .value(chVol) = 0.1
        .value(chVoh) = 0.8     '1.8
        .value(chIol) = 0.001
        .value(chIoh) = -0.001
        .value(chVt) = 1
        .value(chVcl) = -1
        .value(chVch) = 3
        If dibonTesterType = "Jaguar" Then
            .DriverMode = tlDriverModeLargeHiZ  'Largeswing-HiZ
        ElseIf dibonTesterType = "" Then
            .DriverMode = tlDriverModeHiZ  'Largeswing-HiZ
        End If
        TheHdw.Wait 0.01
    End With

    TheHdw.Digital.Timing.period("DIBC_MeasFreq").value = 1 / (1 * 1000000)
    TheHdw.Digital.pins(strSourePinlist).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (1 * 1000000) / 2

    plstTestPinlist.value = strTestPinlist
    Call FreqSweptTest_1MHz(plstTestPinlist, 0.01, patTestPattern, 1000000, -0.2, 1.8, 0.01, tNum, 20)
    TheHdw.Digital.pins(Loop_Back_pins).Disconnect
'
    '''reset
    TheHdw.Wait 0.01
    TheHdw.Digital.ApplyLevelsTiming False, False, False
    TheHdw.Wait 0.01
    TheHdw.Digital.pins(Loop_Back_pins).Disconnect
    If TheExec.sites.ActiveCount = 0 Then Exit Function

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function FreqSweptTest_1MHz(meas_pin As PinList, MeasInterval As Double, pat As Pattern, MinFreq As Double, StartVolt As Double, StopVolt As Double, Resolution As Double, tNum As String, Optional MaxCheckTime As Integer = 20)
On Error GoTo errHandler
    Dim LowerVoh As New PinListData
    Dim UpperVoh As New PinListData
    Dim VohSwept As New PinListData
    Dim MidVolt As Double
    MidVolt = StartVolt + (StopVolt - StartVolt) / 2
    TheHdw.patterns(pat).Load

    MeasFreq_VohSwept meas_pin.value, MeasInterval, StartVolt, MidVolt, Resolution, MinFreq, LowerVoh, pat, MaxCheckTime
    MeasFreq_VohSwept meas_pin.value, MeasInterval, MidVolt, StopVolt, Resolution, MinFreq, UpperVoh, pat, MaxCheckTime
    VohSwept = UpperVoh.Math.Subtract(LowerVoh)

    ' print datalog
    Dim LowLimit As Double: LowLimit = 0.9
    Dim hiLimit As Double: hiLimit = 1.1
    Dim pldMeasureValue As New PinListData  'mask it and change to return measure value as below 20201102
    Set pldMeasureValue = VohSwept          'mask it and change to return measure value as below 20201102
    Util_Judge_Result pldMeasureValue, tNum, "LOOP_BACK_1MHz_VOH", "V"

    If TheExec.sites.ActiveCount = 0 Then Exit Function
    If TheExec.TesterMode = testModeOnline Then
        TheHdw.Digital.Patgen.Continue 0, cpuA
        If (TheHdw.Digital.Patgen.IsRunning = True) Then
            TheHdw.Digital.Patgen.Halt
        End If
    End If
    Exit Function

errHandler:
    MsgBox "Error encountered in FreqTest " + vbCrLf + "VBT Error # " + Trim(str(err.number)) + ": " + err.Description
End Function

Private Function MeasFreq_VohSwept(meas_pin As String, MeasInterval As Double, StartPoint As Double, StopPoint As Double, _
                                   Resolution As Double, MinFreq As Double, MeasData As PinListData, pat As Pattern, _
                                   Optional MaxCheckTime As Integer = 20)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "MeasFreq_VohSwept"

    Dim LowerPointData As ExecData
    Dim UpperPointData As ExecData
    Dim FirstPointData As ExecData
    Dim NextVohPoint As New PinListData
    Dim site As Variant
    Dim stopwhile As Boolean
    Dim mExecResult() As ExecData
    Dim nExecData As ExecData
    Dim Pin_Arr1() As String
    Dim PinCount As Long
    Dim ExecNum As Integer
    Dim PinNum As Integer
    Dim i As Integer
    Dim PrintDebug As Boolean:: PrintDebug = False 'true
    ReDim mExecResult(MaxCheckTime - 1)
    TheExec.DataManager.DecomposePinList meas_pin, Pin_Arr1(), PinCount
    InitPinListData Pin_Arr1, MeasData
    InitExecData Pin_Arr1, LowerPointData
    InitExecData Pin_Arr1, UpperPointData
    InitExecData Pin_Arr1, FirstPointData
    For i = 0 To UBound(mExecResult)
       InitExecData Pin_Arr1, mExecResult(i)
    Next i

    ' init data
    For PinNum = 0 To UBound(LowerPointData.CurrMeasData)
        For Each site In TheExec.sites.Active
           LowerPointData.CurrMeasData(PinNum).Threshold(site) = StartPoint
           UpperPointData.CurrMeasData(PinNum).Threshold(site) = StopPoint
           FirstPointData.CurrMeasData(PinNum).Threshold(site) = (LowerPointData.CurrMeasData(PinNum).Threshold(site) + _
           UpperPointData.CurrMeasData(PinNum).Threshold(site)) / 2
        Next site
    Next PinNum

    Call MeasFreqFunc(meas_pin, MeasInterval, LowerPointData, MinFreq, pat, 100)
    Call MeasFreqFunc(meas_pin, MeasInterval, UpperPointData, MinFreq, pat, 0)
    Call MeasFreqFunc(meas_pin, MeasInterval, FirstPointData, MinFreq, pat, 50)

    mExecResult(0).PrevLowerData = LowerPointData.CurrMeasData
    mExecResult(0).PrevUpperData = UpperPointData.CurrMeasData
    mExecResult(0).CurrMeasData = FirstPointData.CurrMeasData
    If (PrintDebug = True) Then
      PrintDebugLog mExecResult(0), 0
    End If
    stopwhile = CheckStopCondition(mExecResult(0))
    ExecNum = 1
    While (stopwhile = False And ExecNum < MaxCheckTime)
       FindNextExecDef mExecResult(ExecNum), mExecResult(ExecNum - 1), Resolution
       Call MeasFreqFunc(meas_pin, MeasInterval, mExecResult(ExecNum), MinFreq, pat, 50)
       stopwhile = CheckStopCondition(mExecResult(ExecNum))
       If (PrintDebug = True) Then
         PrintDebugLog mExecResult(ExecNum), ExecNum
       End If

       If (stopwhile = False) Then
          ExecNum = ExecNum + 1
       End If
    Wend

    If TheExec.TesterMode = testModeOnline Then 'AL20181130, add only execute online, offline will cause errors.
        ReDim Preserve mExecResult(ExecNum)  '''bug found by rita on 2018Nov08
    End If
    Dim FinalExecData As ExecData
    FinalExecData = mExecResult(UBound(mExecResult)) ' if occur error, please extended StartPoint&StopPoint voltage

    For PinNum = 0 To UBound(FinalExecData.CurrMeasData)
      For Each site In TheExec.sites.Active
       If (FinalExecData.CurrMeasData(PinNum).FinishSearch(site) = True) Then
         MeasData.pins(FinalExecData.CurrMeasData(PinNum).PinName).value(site) = FinalExecData.CurrMeasData(PinNum).Threshold(site)
       Else
         MeasData.pins(FinalExecData.CurrMeasData(PinNum).PinName).value(site) = -1
       End If
      Next site
    Next PinNum

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function MeasFreqFunc(mea_pin As String, MeasInterval As Double, MeasResult As ExecData, MinFreq As Double, pat As Pattern, Optional OfflineData As Integer = 100) As Long
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "MeasFreqFunc"

    Dim freqCnt1 As New PinListData
    Dim PinFreq As New PinListData
    Dim site As Variant
    Dim PinNum As Integer

    For PinNum = 0 To UBound(MeasResult.CurrMeasData)
        For Each site In TheExec.sites.Active
            TheHdw.Digital.pins(MeasResult.CurrMeasData(PinNum).PinName).Levels.value(chVoh) = MeasResult.CurrMeasData(PinNum).Threshold(site)
            TheHdw.Digital.pins(MeasResult.CurrMeasData(PinNum).PinName).Levels.value(chVol) = MeasResult.CurrMeasData(PinNum).Threshold(site)
        Next site
    Next PinNum

    TheHdw.patterns(pat).start

    If (TheExec.TesterMode = testModeOnline) Then
       TheHdw.Digital.Patgen.FlagWait cpuA, 0
    End If

    With TheHdw.Digital.pins(mea_pin).FreqCtr
        .Clear
        .EventSource = VOH 'sweep the level
        .EventSlope = Positive
        .Interval = MeasInterval
        .Enable = IntervalEnable
    End With
    TheHdw.Digital.pins(mea_pin).FreqCtr.start
    freqCnt1 = TheHdw.Digital.pins(mea_pin).FreqCtr.Read

    ' offline data
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)
    TheHdw.Digital.Patgen.HaltWait
    If TheExec.TesterMode = testModeOffline Then FillSimulateData freqCnt1, OfflineData

    PinFreq = freqCnt1.Math.divide(MeasInterval)
    For PinNum = 0 To UBound(MeasResult.CurrMeasData)
       For Each site In TheExec.sites.Active
           MeasResult.CurrMeasData(PinNum).MeasVal = PinFreq.pins(MeasResult.CurrMeasData(PinNum).PinName).value(site)
           If (MeasResult.CurrMeasData(PinNum).MeasVal(site) >= MinFreq) Then
               MeasResult.CurrMeasData(PinNum).MeasPF(site) = 2 'pass, measured expected freq 1MHZ
           Else
               MeasResult.CurrMeasData(PinNum).MeasPF(site) = 1 'fail, not measured the expected freq 1MHZ
           End If
       Next site
    Next PinNum
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function CheckStopCondition(nExecData As ExecData) As Boolean
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "CheckStopCondition"

    CheckStopCondition = False

    Dim PinNum As Integer
    Dim site As Variant
    Dim StopCheck As Boolean: StopCheck = False

    For PinNum = 0 To UBound(nExecData.CurrMeasData)
        For Each site In TheExec.sites.Active
            If nExecData.CurrMeasData(PinNum).FinishSearch(site) = False Then StopCheck = True
        Next site

        If StopCheck = True Then Exit For
    Next PinNum

    CheckStopCondition = Not StopCheck

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function InitPinListData(pins() As String, PinsData As PinListData)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "InitPinListData"
    Dim PinNum As Integer
    Dim site As Variant
    For PinNum = 0 To UBound(pins)
        PinsData.AddPin (pins(PinNum))
        For Each site In TheExec.sites.Active
            PinsData.pins(PinNum).value(site) = -1
        Next site
    Next PinNum
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function InitExecData(pins() As String, nExecData As ExecData)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "InitExecData"
    Dim Pin_Arr1() As String
    Dim PinCount As Long
    Dim PinNum As Integer
    Dim site As Variant

    ReDim nExecData.CurrMeasData(UBound(pins))
    ReDim nExecData.PrevLowerData(UBound(pins))
    ReDim nExecData.PrevUpperData(UBound(pins))

    For PinNum = 0 To UBound(pins)
        nExecData.CurrMeasData(PinNum).PinName = pins(PinNum)
        nExecData.PrevLowerData(PinNum).PinName = pins(PinNum)
        nExecData.PrevUpperData(PinNum).PinName = pins(PinNum)
        nExecData.CurrMeasData(PinNum).FinishSearch = False
        nExecData.CurrMeasData(PinNum).MeasPF = 0
        nExecData.CurrMeasData(PinNum).MeasVal = 0
        nExecData.CurrMeasData(PinNum).Threshold = 0

        nExecData.PrevLowerData(PinNum).FinishSearch = False
        nExecData.PrevLowerData(PinNum).MeasPF = 0
        nExecData.PrevLowerData(PinNum).MeasVal = 0
        nExecData.PrevLowerData(PinNum).Threshold = 0

        nExecData.PrevUpperData(PinNum).FinishSearch = False
        nExecData.PrevUpperData(PinNum).MeasPF = 0
        nExecData.PrevUpperData(PinNum).MeasVal = 0
        nExecData.PrevUpperData(PinNum).Threshold = 0
  Next PinNum
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function PrintDebugLog(MeasData As ExecData, ExecCount As Integer)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "PrintDebugLog"
  Dim PinNum As Integer
  Dim site As Variant
  Dim PrintStr As String
  TheExec.Datalog.WriteComment (" ========== " + CStr(ExecCount) + " ========== ")
    PrintStr = ExtendStr("PinName") + ExtendStr("Type") + ExtendStr("Site", 5) + _
    ExtendStr("Threshold") + ExtendStr("MeasVal") + _
    ExtendStr("P/F") + ExtendStr("Finish")
    TheExec.Datalog.WriteComment (PrintStr)

  For PinNum = 0 To UBound(MeasData.CurrMeasData)
    For Each site In TheExec.sites.Active
       PrintStr = ExtendStr(MeasData.CurrMeasData(PinNum).PinName) + ExtendStr("Current") + ExtendStr(CStr(site), 5) + _
       ExtendStr(CStr(MeasData.CurrMeasData(PinNum).Threshold(site))) + ExtendStr(CStr(MeasData.CurrMeasData(PinNum).MeasVal(site))) + _
       ExtendStr(CStr(MeasData.CurrMeasData(PinNum).MeasPF(site))) + ExtendStr(CStr(MeasData.CurrMeasData(PinNum).FinishSearch(site)))
       TheExec.Datalog.WriteComment (PrintStr)

       PrintStr = ExtendStr(MeasData.PrevLowerData(PinNum).PinName) + ExtendStr("PrevLowerData") + ExtendStr(CStr(site), 5) + _
       ExtendStr(CStr(MeasData.PrevLowerData(PinNum).Threshold(site))) + ExtendStr(CStr(MeasData.PrevLowerData(PinNum).MeasVal(site))) + _
       ExtendStr(CStr(MeasData.PrevLowerData(PinNum).MeasPF(site))) + ExtendStr(CStr(MeasData.PrevLowerData(PinNum).FinishSearch(site)))
       TheExec.Datalog.WriteComment (PrintStr)

        PrintStr = ExtendStr(MeasData.PrevUpperData(PinNum).PinName) + ExtendStr("PrevUpperData") + ExtendStr(CStr(site), 5) + _
       ExtendStr(CStr(MeasData.PrevUpperData(PinNum).Threshold(site))) + ExtendStr(CStr(MeasData.PrevUpperData(PinNum).MeasVal(site))) + _
       ExtendStr(CStr(MeasData.PrevUpperData(PinNum).MeasPF(site))) + ExtendStr(CStr(MeasData.PrevUpperData(PinNum).FinishSearch(site)))
       TheExec.Datalog.WriteComment (PrintStr)
    Next site
  Next PinNum
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function ExtendStr(data As String, Optional exp_length As Integer = 20) As String
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "ExtendStr"
   Dim strLen As Integer
   Dim i As Integer
   Dim NewStr As String
   strLen = Len(data)
   If (strLen < exp_length) Then
     For i = 0 To exp_length - strLen
       NewStr = NewStr + " "
     Next i
     NewStr = NewStr + data
   Else
     NewStr = data
   End If
   ExtendStr = NewStr
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function FindNextExecDef(nExecData As ExecData, PrevExec As ExecData, Resolution As Double)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "FindNextExecDef"
    Dim PinNum As Integer
    Dim site As Variant
    For PinNum = 0 To UBound(PrevExec.CurrMeasData)
      For Each site In TheExec.sites.Active
        If (PrevExec.CurrMeasData(PinNum).FinishSearch(site) = True) Then
           AssignPinMeasData nExecData.PrevLowerData(PinNum), PrevExec.PrevLowerData(PinNum), site
           AssignPinMeasData nExecData.PrevUpperData(PinNum), PrevExec.PrevUpperData(PinNum), site
           AssignPinMeasData nExecData.CurrMeasData(PinNum), PrevExec.CurrMeasData(PinNum), site
        Else
            If (PrevExec.CurrMeasData(PinNum).MeasPF(site) <> PrevExec.PrevLowerData(PinNum).MeasPF(site)) Then
                AssignPinMeasData nExecData.PrevUpperData(PinNum), PrevExec.CurrMeasData(PinNum), site
                AssignPinMeasData nExecData.PrevLowerData(PinNum), PrevExec.PrevLowerData(PinNum), site
            ElseIf (PrevExec.CurrMeasData(PinNum).MeasPF(site) <> PrevExec.PrevUpperData(PinNum).MeasPF(site)) Then
                AssignPinMeasData nExecData.PrevUpperData(PinNum), PrevExec.PrevUpperData(PinNum), site
                AssignPinMeasData nExecData.PrevLowerData(PinNum), PrevExec.CurrMeasData(PinNum), site
            Else
                AssignPinMeasData nExecData.PrevUpperData(PinNum), PrevExec.PrevUpperData(PinNum), site
                AssignPinMeasData nExecData.PrevLowerData(PinNum), PrevExec.PrevLowerData(PinNum), site
                AssignPinMeasData nExecData.CurrMeasData(PinNum), PrevExec.CurrMeasData(PinNum), site
            End If

            If (Math.Abs((nExecData.PrevLowerData(PinNum).Threshold(site) - nExecData.PrevUpperData(PinNum).Threshold(site))) > Resolution) Then
                nExecData.CurrMeasData(PinNum).Threshold(site) = (nExecData.PrevLowerData(PinNum).Threshold(site) + _
                nExecData.PrevUpperData(PinNum).Threshold(site)) / 2
            Else
                AssignPinMeasData nExecData.CurrMeasData(PinNum), PrevExec.CurrMeasData(PinNum), site
                nExecData.CurrMeasData(PinNum).FinishSearch(site) = True

            End If
        End If
      Next site
    Next PinNum
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function FillSimulateData(MeasData As PinListData, OfflineData As Integer)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "FillSimulateData"
    Dim PinNum As Integer
    Dim rndvalue As Integer
    Dim site As Variant
    For PinNum = 0 To MeasData.pins.Count - 1
        For Each site In TheExec.sites.Active
            rndvalue = Int(((OfflineData + 10) - OfflineData) * Rnd + OfflineData)
            MeasData.pins(PinNum).value(site) = rndvalue
        Next site
    Next PinNum
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function AssignPinMeasData(TargetData As PinMeasData, SrcData As PinMeasData, site As Variant)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "AssignPinMeasData"
    TargetData.Threshold(site) = SrcData.Threshold(site)
    TargetData.FinishSearch(site) = SrcData.FinishSearch(site)
    TargetData.MeasPF(site) = SrcData.MeasPF(site)
    TargetData.MeasVal(site) = SrcData.MeasVal(site)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function DIBC_Lib_LoopBack_Freq(tNum As String, Loop_Back_pins_ary As Variant)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DIBC_Lib_LoopBack_Freq"
'    On Error GoTo errHandler
    Dim CountVal As New PinListData
    Dim VihVal As Double
    Dim VodVal As Double
    Dim site As Variant
    Dim pathPathStr_freq As String
    Dim rx_p As String
    Dim rx_n As String
    Dim tx_diff As String
    Dim rx_pins As String
    Dim tx_pins As String
    Dim Lo_Limit As Double
    Dim Hi_Limit As Double
    Dim Tname As String

    Dim LoopBackPinsStr As String
    rx_n = Loop_Back_pins_ary(0)
    tx_diff = Loop_Back_pins_ary(1) 'Loop_Back_pins(2) + "," + Loop_Back_pins(3)
    LoopBackPinsStr = Join(Loop_Back_pins_ary, ",")
    If dibonTesterType = "Jaguar" Then
        pathPathStr_freq = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\" + "DIBC_loopFreqMeas.pat"
    ElseIf dibonTesterType = "UltraFLEXplus" Then
        pathPathStr_freq = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\" + "DIBC_loopFreqMeas.patx"
    End If

    TheHdw.patterns(pathPathStr_freq).Load

    TheHdw.Digital.pins(LoopBackPinsStr).Connect
    TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, PinLevelsSheet:="DIBC_Levels__FT", TimeSetSheet:="DIBC_TSets__FT"
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chVil) = 0
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chVih) = 1.5 'Sicily update from 1v to 1.5v
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chVol) = 0.4
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chVoh) = 0.4
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chIol) = 0
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chIoh) = 0
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chVt) = 0.4
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chVcl) = -1
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.value(chVch) = 6
    TheHdw.Digital.pins(LoopBackPinsStr).Levels.DriverMode = tlDriverModeVt
    TheHdw.Wait 0.01

    ' 62.5MHZ
    TheHdw.Digital.Timing.period("DIBC_MeasFreq_Loopback").value = 1 / (62.5 * 1000000)
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_Loopback", chEdgeD1) = 1 / (62.5 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_Loopback", chEdgeD2) = 1 / (62.5 * 1000000)
    TheHdw.Digital.Timing.ApplyTimingChanges True
    TheHdw.patterns(pathPathStr_freq).start
    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    TheHdw.Wait 0.1

    With TheHdw.Digital.pins(tx_diff).FreqCtr 'tx_diff
        .Clear
        .EventSlope = Positive
        .EventSource = vol
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With

    CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    CountVal = CountVal.Math.divide(TheHdw.Digital.pins(tx_diff).FreqCtr.Interval)

    Call TheHdw.Digital.Patgen.Continue(0, cpuA)

    TheHdw.Digital.Patgen.HaltWait

    Dim li As Long
    Dim vsite As Variant
    Dim pldPosMeas As New PinListData
    pldPosMeas.AddPin CountVal.pins(0).name
    For Each vsite In TheExec.sites.Active
        pldPosMeas = CountVal.pins(0).value(vsite)
    Next vsite

    Util_Judge_Result pldPosMeas, tNum, "LOOP_BACK_Freq", "Hz"
    tNum = CStr(CLng(tNum) + 1)

    If TheExec.sites.ActiveCount = 0 Then Exit Function

    TheHdw.Digital.Timing.period("DIBC_MeasFreq_Loopback").value = 1 / (125 * 1000000)
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_Loopback", chEdgeD1) = 1 / (125 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_Loopback", chEdgeD2) = 1 / (125 * 1000000)
    TheHdw.Digital.Timing.ApplyTimingChanges True
    TheHdw.patterns(pathPathStr_freq).start
    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    TheHdw.Wait 0.1

    With TheHdw.Digital.pins(tx_diff).FreqCtr
        .Clear
        .EventSlope = Positive
        .EventSource = vol
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With

    CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    CountVal = CountVal.Math.divide(TheHdw.Digital.pins(tx_diff).FreqCtr.Interval)

    Call TheHdw.Digital.Patgen.Continue(0, cpuA)
    TheHdw.Digital.Patgen.HaltWait

    Set pldPosMeas = Nothing 'clean data

    pldPosMeas.AddPin CountVal.pins(0).name
    For Each vsite In TheExec.sites.Active
        pldPosMeas = CountVal.pins(0).value(vsite)
    Next vsite
    Util_Judge_Result pldPosMeas, CLng(tNum) + 1, "LOOP_BACK_Freq", "Hz"
    tNum = CStr(CLng(tNum) + 1)

    If TheExec.sites.ActiveCount = 0 Then Exit Function

    ' Dim Loop_count As Integer '---WY_MOD
    ' Loop_count = 0

' Re_meas_F:
    ' TheHdw.Wait 0.1

    '250MHZ
    ' TheHdw.Digital.Timing.period("DIBC_MeasFreq_Loopback").Value = 1 / (200 * 1000000)
    ' TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_Loopback", chEdgeD1) = 1 / (200 * 1000000) / 2
    ' TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_Loopback", chEdgeD2) = 1 / (200 * 1000000)
    ' TheHdw.Digital.Timing.ApplyTimingChanges True
    ' TheHdw.Patterns(pathPathStr_freq).start
    ' Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    ' TheHdw.Wait 0.2

    ' With TheHdw.Digital.pins(tx_diff).FreqCtr
        ' .Clear
        ' .EventSlope = Positive
        ' .EventSource = vol
        ' .Interval = 0.01
        ' .Enable = IntervalEnable
        ' .start
    ' End With

    ' CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    ' CountVal = CountVal.Math.Divide(TheHdw.Digital.pins(tx_diff).FreqCtr.Interval)

    ' Call TheHdw.Digital.Patgen.Continue(0, cpuA)

    ' TheHdw.Digital.Patgen.HaltWait
'================loop retest===================='---WY_MOD
    ' For Each vsite In TheExec.sites.Active
      ' If (CountVal.pins(li).value(vsite) > 218000000 Or CountVal.pins(li).value(vsite) < 190000000) And Loop_count < 5 Then
        ' Loop_count = Loop_count + 1
        ' GoTo Re_meas_F
      ' End If
    ' Next vsite
'================loop retest====================

    ' Set pldPosMeas = Nothing 'clean data

    ' pldPosMeas.AddPin CountVal.pins(li).Name
    ' For Each vsite In TheExec.sites.Active
        ' pldPosMeas = CountVal.pins(li).Value(vsite)
    ' Next vsite

    ' Util_Judge_Result pldPosMeas, CLng(tNum) + 2, "LOOP_BACK_Freq", "Hz"
    ' If TheExec.sites.ActiveCount = 0 Then Exit Function
    TheHdw.Digital.pins(LoopBackPinsStr).Disconnect

'    Exit Function


Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function DIBC_Lib_Loop_Back_MCC_Cap_meas(TestNum As String, Loop_Back_pins As Variant, dblTestMeasureCurrent As Double)
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "DIBC_Lib_Loop_Back_MCC_Cap_meas"

    Dim RX_Pinname As String: RX_Pinname = Loop_Back_pins(0)
    Dim TX_Pinname As String: TX_Pinname = Loop_Back_pins(1)
    Dim pat        As String:

    ' connect PPMU'
    TheHdw.Digital.pins(RX_Pinname).Disconnect
    TheHdw.Digital.pins(TX_Pinname).Disconnect
    TheHdw.PPMU.pins(RX_Pinname).Connect
    TheHdw.PPMU.pins(TX_Pinname).Connect
    TheHdw.PPMU.pins(RX_Pinname).ForceV 0
    TheHdw.PPMU.pins(RX_Pinname).Gate = tlOn
    TheHdw.Wait 0.004
    If dibonTesterType = "Jaguar" Then
        pat = "CX_HSD_" + TX_Pinname + ".pat" '"LoopBack_Cap_Meas.pat"
    ElseIf dibonTesterType = "UltraFLEXplus" Then
        pat = "CX_HSD_" + TX_Pinname + ".patx" '"LoopBack_Cap_Meas.patx"
    End If
    ' capacitor test'
    Dim Tname As String: Tname = Replace(TheExec.DataManager.instancename, "_cap_value", "")
    Dim strHSD_Pins As String: strHSD_Pins = TX_Pinname
    Dim vsite As Variant
    Dim dblPeriod As Double: dblPeriod = 20 * ns
    Dim pldMeasureValue As New PinListData
    Dim dblTestChargeVoltage As Double:     dblTestChargeVoltage = 0.5
    Dim dblTestOffsetVoltage As Double:     dblTestOffsetVoltage = 0.3
    Dim pldFailCount As New PinListData
    Dim strPatternStartLabel As String:     strPatternStartLabel = "500usPer020ns" 'msPer020ns"'"003msPer020ns"
    Dim strTestPattern As String
    Dim strPatternPath As String
    Dim Label_str As String

    strTestPattern = pat ' "CX_HSD_PCIE_UP_TX0_N.pat"
    strPatternPath = TheExec.TestProgram.Path + "\\" + DIBC_FOLDER + "\\Patterns\\"

    With TheHdw.PPMU.pins(strHSD_Pins)
        .ForceV 0
        .Gate = tlOn
        TheHdw.Wait 0.005
        .Gate = tlOff
        TheHdw.Wait 0.002
        .Disconnect
    End With

    TheHdw.Digital.ApplyLevelsTiming False, True, True, , , , strHSD_Pins

    If TheHdw.tester.type = "Jaguar" Then 'uflx 20201102
        TheHdw.Digital.ConnectPins (strHSD_Pins)
        TheHdw.Wait 0.002
        With TheHdw.Digital.pins(strHSD_Pins).Levels
            .value(chVih) = 2
            .value(chVil) = -0.5
            .value(chVol) = dblTestOffsetVoltage
            .value(chVch) = 5.5
            .value(chVcl) = -1
            .value(chVoh) = dblTestOffsetVoltage + dblTestChargeVoltage
            .value(chIoh) = 0
            .value(chIol) = dblTestMeasureCurrent
            .value(chVt) = 5.5
        End With

        With TheHdw.PinLevels.pins(strHSD_Pins) 'mask it 20201102
            .ModifyLevel chVih, 2
            .ModifyLevel chVil, -0.5
            .ModifyLevel chVol, dblTestOffsetVoltage
            .ModifyLevel chVch, 5.5
            .ModifyLevel chVcl, -1
        End With

        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVoh, dblTestOffsetVoltage + dblTestChargeVoltage
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIoh, 0
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIol, dblTestMeasureCurrent
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVt, 5.5
        TheHdw.Digital.ConnectPins (strHSD_Pins) 'uflx 20201102

        If CInt(left(TheExec.SoftwareVersion, 1)) < 8 Then
            TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Load 'uflx 20201102
        End If
        TheHdw.Wait 0.002

        'If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
        '    TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Load ''uflx 20201102
        'End If
        TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Load 'uflx 20201102

        Label_str = TX_Pinname + "_CX_" ' "LoopBack_"
        TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Run Label_str & strPatternStartLabel 'uflx 20201102
        TheHdw.Wait 0.005

        pldMeasureValue.AddPin (strHSD_Pins)
        pldFailCount = TheHdw.Digital.pins(strHSD_Pins).FailCount
        TheHdw.Wait 0.005

        For Each vsite In TheExec.sites.Active
            pldMeasureValue.pins(strHSD_Pins).value(vsite) = dblTestMeasureCurrent * ((pldFailCount.pins(strHSD_Pins).value(vsite) * dblPeriod) / dblTestChargeVoltage)
        Next vsite

        With TheHdw.PPMU.pins(strHSD_Pins)
            .ForceV 0
            TheHdw.Wait 0.02
            .Gate = tlOff
            TheHdw.Wait 0.002
        End With

        TheHdw.Digital.DisconnectPins (strHSD_Pins) 'uflx 20201102
        Util_Judge_Result pldMeasureValue, TestNum, "MCC", "F"

        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '''''If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
        '''''    TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Unload 'uflx 20201102
        '''''End If
        'TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Unload 'uflx 20201102

    ElseIf TheHdw.tester.type = "UltraFLEXplus" Then
        TheHdw.Digital.pins(strHSD_Pins).Connect
        TheHdw.Wait 0.002
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chVih) = 2
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chVil) = -0.5
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chVol) = dblTestOffsetVoltage
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chVch) = 5.5
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chVcl) = -1
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chVoh) = dblTestOffsetVoltage + dblTestChargeVoltage
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chIoh) = 0
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chIol) = dblTestMeasureCurrent
        TheHdw.Digital.pins(strHSD_Pins).Levels.value(chVt) = 5.5

        TheHdw.Digital.pins(strHSD_Pins).Connect
        If CInt(left(TheExec.SoftwareVersion, 1)) < 8 Then
            TheHdw.patterns(strPatternPath & strTestPattern).Load
        End If
        TheHdw.Wait 0.002

        'If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
        '    TheHdw.Patterns(strPatternPath & strTestPattern).Load
        'End If
        'TheHdw.Patterns(strPatternPath & strTestPattern).Load

        Label_str = TX_Pinname & "_CX_"

        'If Label_str Like "GP_*" Or Label_str Like "ST_*" Then  ' For Sicily 20200326
        '   Label_str = Mid(Label_str, 4)
        'End If

        TheHdw.patterns(strPatternPath & strTestPattern).start ((Label_str & strPatternStartLabel))
        TheHdw.Digital.Patgen.HaltWait
        TheHdw.Wait 0.005

        pldMeasureValue.AddPin (strHSD_Pins)
        pldFailCount = TheHdw.Digital.pins(strHSD_Pins).FailCount
        TheHdw.Wait 0.005

        For Each vsite In TheExec.sites.Active
            pldMeasureValue.pins(strHSD_Pins).value(vsite) = dblTestMeasureCurrent * ((pldFailCount.pins(strHSD_Pins).value(vsite) * dblPeriod) / dblTestChargeVoltage)
        Next vsite

        With TheHdw.PPMU.pins(strHSD_Pins)
            .ForceV 0
            TheHdw.Wait 0.02
            .Gate = tlOff
            TheHdw.Wait 0.002
        End With

        TheHdw.Digital.pins(strHSD_Pins).Disconnect

        Util_Judge_Result pldMeasureValue, TestNum, "MCC", "F"
        If TheExec.sites.ActiveCount = 0 Then Exit Function

        'If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
        '    TheHdw.Patterns(strPatternPath & strTestPattern).Unload
        'End If
        'TheHdw.Patterns(strPatternPath & strTestPattern).Unload

    End If

    Set pldMeasureValue = Nothing

    ' disconnect PPMU
    TheHdw.PPMU.pins(RX_Pinname).Gate = tlOff
    TheHdw.PPMU.pins(TX_Pinname).Gate = tlOff
    TheHdw.Wait 0.002
    TheHdw.PPMU.pins(RX_Pinname).Reset tlResetSettings
    TheHdw.PPMU.pins(TX_Pinname).Reset tlResetSettings
    TheHdw.PPMU.pins(RX_Pinname).Disconnect
    TheHdw.PPMU.pins(TX_Pinname).Disconnect

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

Private Function Trace_Resistance_Meas(tNum As String, TestPin As Variant, RelayOn_Pin As String) As Long
On Error GoTo errHandler
    Dim sCurrentFuncName As String:: sCurrentFuncName = "Trace_Resistance_Meas"

Dim MeasureValue As New PinListData
Dim ForceValI As Double
Dim MeasurePinStr As String

Dim PinArr() As String
Dim PinCount As Long
Dim pin  As Variant
Dim Res_Value As New SiteDouble
Dim MeasRes_result As PinListData
Dim i As Integer
Dim hiLimit As Double
Dim LowLimit As Double
Dim force_i As Double
Dim Tname As String
Dim sPin As String
'On Error GoTo errHandler
hiLimit = 10#
LowLimit = 0

MeasurePinStr = TestPin
Tname = "Digital Path Resistance"
''''' thehdw.DCVS.Pins("All_power").Connect tlDCVSConnectDefault
'''''                     thehdw.Wait 0.01
'''''                     With thehdw.DCVS.Pins("All_power")
'''''                        .mode = tlDCVSModeVoltage
'''''                        .Voltage.Main.Value = 0#
'''''                        .Voltage.Alt.Value = 0#
'''''                        .Voltage.output = tlDCVSVoltageMain
'''''                        .CurrentRange.Value = 0.02
'''''                        .Meter.mode = tlDCVSMeterCurrent
'''''                        .Meter.CurrentRange = 0.02
'''''                        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
'''''                        .Gate = True
'''''                        thehdw.Wait 1#
'''''                      End With
Dim Relay_on_pin_temp1() As String 'add 191121
Dim Relay_on_pin_temp2() As String 'add 191121
Dim Relay_Name As String
If RelayOn_Pin <> "" Then
    Relay_on_pin_temp1 = Split(RelayOn_Pin, ";") 'add 191121
    For i = 0 To UBound(Relay_on_pin_temp1)
        Relay_on_pin_temp2 = Split(Relay_on_pin_temp1(i), "+")
        Relay_On_Pin_Dic_short(Relay_on_pin_temp2(0)) = Relay_on_pin_temp2(1)
    Next i
End If
Dim Current_Range As Double
'TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
PinArr = TestPin
For i = 0 To UBound(PinArr)
'For Each pin In PinArr
    MeasureValue.AddPin (PinArr(i))
    'For Each Site In TheExec.sites.Selected
             'theexec.Sites(site).selectd
           sPin = PinArr(i)

         If Relay_On_Pin_Dic_short.Exists(PinArr(i)) Then
                Relay_Name = Relay_On_Pin_Dic_short(PinArr(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOn
                TheHdw.Wait 0.0001

          End If


      TheHdw.Digital.pins(PinArr(i)).Disconnect
      With TheHdw.PPMU(PinArr(i))

          .Connect

          .ForceV 0.1, 0.04
          .Gate = tlOn
           TheHdw.Wait 0.01
           'DebugPrintFunc_PPMU leakage_pins.Value
           MeasureValue = .Read(tlPPMUReadMeasurements)
          ' current_range = thehdw.PPMU.Pins(PinArr(i)).MeasureCurrentRange
            Res_Value = TheHdw.PPMU.pins(PinArr(i)).Voltage.value / MeasureValue.pins(PinArr(i)).value
          .ForceV (0)
          .Gate = tlOff
          .Disconnect
      End With
         '---change judge hi/li limit here for different spec 20190807----
          Dim sPin_Key As String
          Dim bb As Integer
         ' Dim Site As Variant
         ' Dim MeasureValue_site() As New SiteDouble

          'MeasureValue_site = pldMeasureValue
         ' For Each Site In theexec.sites.Existing
         If PinArr(i) = "VDD_CPU_Monitor" Then
            PinArr(i) = UCase("VDD_CPU_Monitor")
         End If


''''                     sPin_Key = CStr(Site) & CStr(PinArr(i)) 'no path hi/li limit sheet
''''                     HiLimit = Dict_Res_Hlimit(sPin_Key)
                     MeasRes_result.AddPin (PinArr(i))
                     MeasRes_result.pins(MeasRes_result) = Res_Value

                    ' TheExec.Flow.TestLimit resultVal:=Res_Value, PinName:=PinArr(i), lowVal:=LowLimit, hiVal:=HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=force_i, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone

         ' Next Site
        '---change judge hi/li limit here for different spec 20190807----


            TheHdw.Wait 0.001
            'MeasureValue.Pins(PinArr(i)) = thehdw.PPMU.Pins(pin).Read(tlPPMUReadMeasurements, 10)

           ' thehdw.PPMU.Pins(PinArr(i)).Disconnect


            With TheHdw.PPMU.pins(PinArr(i))

              .Gate = tlOff
              .Disconnect
            End With

         If Relay_On_Pin_Dic_short.Exists(PinArr(i)) Then
                Relay_Name = Relay_On_Pin_Dic_short(PinArr(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOff
                TheHdw.Wait 0.0001

          End If



    'Next Site
 'Next pin
  Next i

  Util_Judge_Result MeasRes_result, tNum, "Path", Ohm
'Exit Function


Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + sCurrentFuncName + ":: please check it out."
End Function

'221012 Ben Modify
Private Function DCVI_Cap_Meas_2(tNum As String, Test_PinName As String, Optional Relay_On As String, Optional sCurrent As Double, Optional sChargeVoltage As Double, Optional sVoltageOffset As Double, Optional Relay_S As String)
        
        On Error GoTo errHandler
        
        ' If Relay_S <> "" Then
            ' If sRelaystatus = "ON" Then
            ' Call SmartRelaySwitch(Relay_on & "," & DGS_Relay & "," & Relay_S)
            ' Else
            ' Call SmartRelaySwitch(Relay_S & "," & DGS_Relay)
            ' End If
        ' Else
            ' If sRelaystatus = "ON" Then
            ' Call SmartRelaySwitch(Relay_on & "," & DGS_Relay)
            ' Else
            ' Call SmartRelaySwitch(DGS_Relay)
            ' End If
        ' End If
                TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
        
        '***connect DCVI instrument***'
        TheHdw.DCVI.pins(Test_PinName).Connect tlDCVIConnectDefault
        TheHdw.Wait 0.0005
        '***-----------------------***'

        '********************capacitor test begin********************'
        Dim vsite As Variant
        Dim intI As Integer
        Dim intJ As Integer
        Dim pldMeasureValue As New PinListData
        Dim v1 As New PinListData
        Dim v2 As New PinListData
        Dim deltaV As New PinListData
        Dim t1 As New PinListData
        Dim t2 As New PinListData
        Dim deltaT As New PinListData
        Dim dT As Double
        Dim DCVIResultDSP As New DSPWave
        Dim strDCVI_PinArray() As String
        Dim dblCapSampleRate As Double
        Dim dblCapSampleSize As Double
        Dim dblCurrentRange As Double
        Dim dblCaptureWait As Double
        Dim strTestDCVI_Pin As String
        Dim dblTestMeasureCurrent As Double
        Dim dblTestChargeVoltage As Double
        Dim dblTestOffsetVoltage As Double
        Dim slIndex1 As New SiteLong
        Dim slIndex2 As New SiteLong
        
        strTestDCVI_Pin = Test_PinName
        dblTestMeasureCurrent = sCurrent
        dblTestChargeVoltage = sChargeVoltage
        dblTestOffsetVoltage = sVoltageOffset


        strDCVI_PinArray = Split(Test_PinName, ",")

        dblCapSampleRate = 100 * kHz
        dblCapSampleSize = 512
        dblCaptureWait = dblCapSampleSize / dblCapSampleRate
        dblCurrentRange = 0.005

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmOff

        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.Add ("CapSignal")
        With TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals("CapSignal")
                .SampleRate = 100000
                .SampleSize = 512
                .LoadSettings
        End With

        '***set DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Current = 0.02
                .Voltage = 5
                .VoltageRange.value = 7
                .CurrentRange.Autorange = True
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 7
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'

        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblTestMeasureCurrent
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait dblCaptureWait
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'computing of the capacitor value
        For intJ = 0 To UBound(strDCVI_PinArray)
                v1.AddPin (strDCVI_PinArray(intJ))
                v2.AddPin (strDCVI_PinArray(intJ))
                t1.AddPin (strDCVI_PinArray(intJ))
                t2.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
'''                                        m_strCapWavePath = ".\" & DIBC_FOLDER & "Patterns\waves\"
'''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
'''                                        Call DCVIResultDSP.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
'''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

        'DCVIResultDSP.Plot

                                slIndex1 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                slIndex2 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                t1.pins(strDCVI_PinArray(intJ)).value = slIndex1 / dblCapSampleRate
                                t2.pins(strDCVI_PinArray(intJ)).value = slIndex2 / dblCapSampleRate
                                If slIndex1 <> -1 Then
                                        v1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(slIndex1)
                                Else
                                        v1.pins(strDCVI_PinArray(intJ)).value = 0
                                End If
                                If slIndex2 <> -1 Then
                                        v2.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(slIndex2)
                                Else
                                        v2.pins(strDCVI_PinArray(intJ)).value = 0
                                End If

                                dT = t2.pins(strDCVI_PinArray(intJ)).value - t1.pins(strDCVI_PinArray(intJ)).value
                                If dT = 0 Then
                                        t2.pins(strDCVI_PinArray(intJ)).value = t2.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        deltaV.AddPin (strTestDCVI_Pin)
        deltaV = v2.Math.Subtract(v1)
        deltaT.AddPin (strTestDCVI_Pin)
        deltaT = t2.Math.Subtract(t1)

        For Each vsite In TheExec.sites.Active
                If TheExec.sites.Active Then
                        If deltaV.pins(strTestDCVI_Pin).value(vsite) = 0 Then
                                deltaV.pins(strTestDCVI_Pin).value(vsite) = 0.000000000001
                        End If
                End If
        Next vsite

        pldMeasureValue.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValue = deltaT.Math.Multiply(dblTestMeasureCurrent)
        pldMeasureValue = pldMeasureValue.Math.divide(deltaV)

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmDefault

        'discharge the cap
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01

        '********************main test begin********************'
        Util_Judge_Result pldMeasureValue, tNum, strTestDCVI_Pin & "," & Relay_On, "F"
           'Util_Judge_Result pldMeasureValue, tNum, strTestDCVI_Pin & "," & Relay_on & " " & sRelaystatus, "F"
        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************main test end********************'

        Set pldMeasureValue = Nothing
        '********************capacitor test end********************'

        '***disconnect DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .Gate = False
                TheHdw.Wait 0.002
                .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 0.0005
        '***-----------------------***'

     'Call SmartRelaySwitch("")

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in DCVI_Cap_meas_2"
End Function



' ********************************************************************************
' =========               Auto Generated DIBC_Test                       =========
' ********************************************************************************

Private Sub FT_DIBC_Test_4000()
    FT_DIBC_Lib_MeasureIOleakage 4000, Array()
End Sub

Private Sub FT_DIBC_Test_4100()
    FT_DIBC_Lib_UtilityCheck 4100, Array()
End Sub

Private Sub FT_DIBC_Test_5000()
    FT_DIBC_Lib_SPI_ROM 5000, Array("K110", "K111", "K58", "K59", "K60", "K61")
End Sub

Private Sub FT_DIBC_Test_10000()
    FT_DIBC_Lib_UP1600_IO 10000, Array("AON_SLEEP1_RESETN", "AOP_I2CM0_SCL", "AOP_I2CM0_SDA", "AOP_I2CM1_SCL", "AOP_I2CM1_SDA", "AOP_I2CM2_SCL", "AOP_I2CM2_SDA", "AOP_I2CM3_SCL", "AOP_I2CM3_SDA", "AOP_I2S0_BCLK", "AOP_I2S0_DIN", "AOP_I2S0_DOUT", "AOP_I2S0_LRCK", "AOP_I2S0_MCK", "AOP_I2S1_BCLK", "AOP_I2S1_DIN", "AOP_I2S1_DOUT", "AOP_I2S1_LRCK", "AOP_I2S1_MCK", "AOP_I2S2_BCLK", "AOP_I2S2_DIN", "AOP_I2S2_DOUT", "AOP_I2S2_LRCK", "AOP_I2S2_MCK", "AOP_I2S3_BCLK", "AOP_I2S3_DIN", "AOP_I2S3_DOUT", "AOP_I2S3_LRCK", "AOP_I2S4_BCLK", "AOP_I2S4_DIN", "AOP_I2S4_DOUT", "AOP_I2S4_LRCK", "AOP_I2S5_BCLK", "AOP_I2S5_DIN", "AOP_I2S5_DOUT", _
"AOP_I2S5_LRCK", "AOP_SCM_IO0", "AOP_SCM_IO1", "AOP_SCM_IO10", "AOP_SCM_IO11", "AOP_SCM_IO12", "AOP_SCM_IO13", "AOP_SCM_IO14", "AOP_SCM_IO15", "AOP_SCM_IO2", "AOP_SCM_IO3", "AOP_SCM_IO4", "AOP_SCM_IO5", "AOP_SCM_IO6", "AOP_SCM_IO7", "AOP_SCM_IO8", "AOP_SCM_IO9", "AOP_SPI0_MISO", "AOP_SPI0_MOSI", "AOP_SPI0_SCLK", "AOP_SPI1_MISO", "AOP_SPI1_MOSI", "AOP_SPI1_SCLK", "AOP_SPMI0_SCLK", "AOP_SPMI0_SDATA", "AOP_SPMI1_SCLK", "AOP_SPMI1_SDATA", "AOP_UART2_RXD", "AOP_UART2_TXD", "ATC0_AUX_N", "ATC0_AUX_P", "ATC0_USB_N", "ATC0_USB_P", "BOARD_ID0", "BOARD_ID1", "BOARD_ID2", "BOARD_ID3", "BOARD_ID4", _
"CFSB", "CFSB_AON", "CFSB_XTAL", "COLD_RESETN", "DBG_PROBE_VALID", "DBG_USB_ED_N", "DBG_USB_ED_P", "DDR0_SYS_ALIVE", "DFU_STATUS", "DISP_EXT_HPD", "DISP_HPD", "DISP_I2C1_SCL", "DISP_I2C1_SDA", "DISP_LINK_PWR_DWN", "DISP_POL", "DISP_TE", "DISP_TOUCH_BSYNC0", "DISP_TOUCH_BSYNC1", "DISP_TOUCH_EB", "DSG_BSYNC0", "DSG_BSYNC1", "DSG_BSYNC2", "DSG_EB", "DSG_PROXBOLED", "FORCE_DFU", "FPWM0", "FPWM1", "FPWM2", "GPIO0", "GPIO1", "GPIO10", "GPIO11", "GPIO12", "GPIO13", "GPIO14", "GPIO15", "GPIO16", "GPIO17", "GPIO18", "GPIO19", "GPIO2", "GPIO20", "GPIO21", "GPIO22", "GPIO23", "GPIO24", "GPIO25", "GPIO3", "GPIO4", _
"GPIO5", "GPIO6", "GPIO7", "GPIO8", "GPIO9", "GP_PCIE_CLKREQ0_L", "GP_PCIE_CLKREQ1_L", "GP_PCIE_PERST0_L", "GP_PCIE_PERST1_L", "GP_PCIE_REF_CLK0_N", "GP_PCIE_REF_CLK0_P", "GP_PCIE_REF_CLK1_N", "GP_PCIE_REF_CLK1_P", "HOLD_RESET", "I2C0_SCL", "I2C0_SDA", "I2C1_SCL", "I2C1_SDA", "I2C2_SCL", "I2C2_SDA", "I2C3_SCL", "I2C3_SDA", "I2C4_SCL", "I2C4_SDA", "ISP_GPIO_0", "ISP_GPIO_1", "ISP_GPIO_2", "ISP_GPIO_3", "ISP_I2C0_SCL", "ISP_I2C0_SDA", "ISP_I2C1_SCL", "ISP_I2C1_SDA", "ISP_SPMI0_SCLK", "ISP_SPMI0_SDATA", "ISP_SPMI1_SCLK", "ISP_SPMI1_SDATA", "JTAG_SEL", "JTAG_TCK", "JTAG_TDI", "JTAG_TDO", "JTAG_TMS", _
"JTAG_TRSTN", "KIS_TO_PMU_REQUEST_DFU", "KIS_TO_PMU_RESET", "LPDPRX_AUX_D0_P", "LPDPRX_AUX_D10_P", "LPDPRX_AUX_D11_P", "LPDPRX_AUX_D12_P", "LPDPRX_AUX_D1_P", "LPDPRX_AUX_D2_P", "LPDPRX_AUX_D3_P", "LPDPRX_AUX_D4_P", "LPDPRX_AUX_D5_P", "LPDPRX_AUX_D6_P", "LPDPRX_AUX_D7_P", "LPDPRX_AUX_D8_P", "LPDPRX_AUX_D9_P", "LPDPTX_AUX_N", "LPDPTX_AUX_P", "LPDPTX_D0_N", "LPDPTX_D0_P", "LPDPTX_D1_N", "LPDPTX_D1_P", "LPDPTX_D2_N", "LPDPTX_D2_P", "LPDPTX_D3_N", "LPDPTX_D3_P", "MIC_DISABLE_L", "MIC_DISABLE_WARN_L", "MIPIC0_CLK_N", "MIPIC0_CLK_P", "MIPIC0_DATA0_N", "MIPIC0_DATA0_P", "MIPIC0_DATA1_N", _
"MIPIC0_DATA1_P", "MIPID_CLKN", "MIPID_CLKP", "MIPID_DATAN0", "MIPID_DATAN1", "MIPID_DATAN2", "MIPID_DATAN3", "MIPID_DATAP0", "MIPID_DATAP1", "MIPID_DATAP2", "MIPID_DATAP3", "MTP_FUNC0", "MTP_FUNC1", "MTP_FUNC2", "MTP_FUNC3", "MTP_FUNC4", "MTP_I2C0_SCL", "MTP_I2C0_SDA", "MTP_SPI0_MISO", "MTP_SPI0_MOSI", "MTP_SPI0_SCLK", "NAND_SYS_CLK", "NUB_CLK_OUT0", "NUB_CLK_OUT1", "NUB_DBGWRAP_TEST", "NUB_DOCK_CONNECT", "NUB_SCM_IO0", "NUB_SCM_IO1", "NUB_SCM_IO10", "NUB_SCM_IO11", "NUB_SCM_IO12", "NUB_SCM_IO13", "NUB_SCM_IO14", "NUB_SCM_IO15", "NUB_SCM_IO2", "NUB_SCM_IO3", "NUB_SCM_IO4", "NUB_SCM_IO5", _
"NUB_SCM_IO6", "NUB_SCM_IO7", "NUB_SCM_IO8", "NUB_SCM_IO9", "NUB_SPMI0_SCLK", "NUB_SPMI0_SDATA", "NUB_SPMI1_SCLK", "NUB_SPMI1_SDATA", "NUB_SWD_TCK_OUT0", "NUB_SWD_TCK_OUT1", "NUB_SWD_TMS0", "NUB_SWD_TMS1", "NUB_SWD_TMS2", "NUB_SWD_TMS3", "NUB_SWD_TMS4", "SENSOR0_CLK", "SENSOR1_CLK", "SGPIO0", "SGPIO1", "SGPIO2", "SI2C0_SCL", "SI2C0_SDA", "SMC_I2CM0_SCL", "SMC_I2CM0_SDA", "SMC_I2CM1_SCL", "SMC_I2CM1_SDA", "SMC_I2CM2_SCL", "SMC_I2CM2_SDA", "SOCHOT1_L", "SOC_ACC_ICTS_AM_EN", "SOC_ACC_ICTS_AM_STR", "SPI0_MISO", "SPI0_MOSI", "SPI0_SCLK", "SPI2_MISO", "SPI2_MOSI", "SPI2_SCLK", "SPI2_SSIN", _
"SPI3_MISO", "SPI3_MOSI", "SPI3_SCLK", "SPI3_SSIN", "SSD_BFH", "SSD_RESET_L", "SS_HOSTMODE", "SS_RST_L", "ST_PCIE_CLKREQ0_L", "ST_PCIE_PERST0_L", "ST_PCIE_REF_CLK0_N", "ST_PCIE_REF_CLK0_P", "SWD_TCK_OUT2", "SWD_TMS5", "SWD_TMS6", "SWD_TMS7", "TESTMODE", "THROTTLE_TRIGGER0_L", "THROTTLE_TRIGGER1_L", "THROTTLE_TRIGGER2_L", "THROTTLE_TRIGGER3_L", "THROTTLE_TRIGGER4_L", "TST_CLKOUT", "UART0_RXD", "UART0_TXD", "UART1_CTSN", "UART1_RTSN", "UART1_RXD", "UART1_TXD", "UART2_CTSN", "UART2_RTSN", "UART2_RXD", "UART2_TXD", "UART3_CTSN", "UART3_RTSN", "UART3_RXD", "UART3_TXD", "UART4_RXD", "UART4_TXD", _
"UID_MODE_L", "UNTRUSTED_BOOTCONFIG", "VSS_1_1", "VSS_2_1", "VSS_3", "VSS_4", "WDOG_L")
End Sub

Private Sub FT_DIBC_Test_40000()
    FT_DIBC_Lib_UP1600_RELAY 40000, Array("RT_CLK32768", "RT_CLK32768_PA", "K1")
End Sub

Private Sub FT_DIBC_Test_60000()
    FT_DIBC_Lib_UP1600_RELAY_R_2 60000, Array("ATC0_USB_RESREF", "K46", "K47")
End Sub

Private Sub FT_DIBC_Test_60100()
    FT_DIBC_Lib_UP1600_RELAY_R_2 60100, Array("DBG_USB_RESREF", "K48", "K49")
End Sub

Private Sub FT_DIBC_Test_60200()
    FT_DIBC_Lib_UP1600_RELAY_R_2 60200, Array("DIE_CORNER_DC_NE", "K56", "K57")
End Sub

Private Sub FT_DIBC_Test_60300()
    FT_DIBC_Lib_UP1600_RELAY_R_2 60300, Array("DIE_CORNER_DC_NW", "K50", "K51")
End Sub

Private Sub FT_DIBC_Test_60400()
    FT_DIBC_Lib_UP1600_RELAY_R_2 60400, Array("DIE_CORNER_DC_SE", "K52", "K53")
End Sub

Private Sub FT_DIBC_Test_60500()
    FT_DIBC_Lib_UP1600_RELAY_R_2 60500, Array("DIE_CORNER_DC_SW", "K54", "K55")
End Sub

Private Sub FT_DIBC_Test_60600()
    FT_DIBC_Lib_UP1600_RELAY_R_2 60600, Array("MIPIC0_REXT", "K42", "K43")
End Sub

Private Sub FT_DIBC_Test_60700()
    FT_DIBC_Lib_UP1600_RELAY_R_2 60700, Array("MIPID_REXT", "K44", "K45")
End Sub

Private Sub FT_DIBC_Test_70000()
    FT_DIBC_Lib_UP1600_RELAY_2 70000, Array("XI0", "XI0_PA", "XO0", "XO0_PA", "K73")
End Sub

Private Sub FT_DIBC_Test_200000()
    FT_DIBC_Lib_UVI80_RELAY 200000, Array("VDD_PCPU_ABSMIN", "K66")
End Sub

Private Sub FT_DIBC_Test_240000()
    FT_DIBC_Lib_UVI80_RELAY_R 240000, Array("ANALOGMUX_OUT_POS_SRC", "K67", "K68", "K69")
End Sub

Private Sub FT_DIBC_Test_240100()
    FT_DIBC_Lib_UVI80_RELAY_R 240100, Array("ANALOGMUX_OUT_NEG_SRC", "K70", "K72", "K71")
End Sub

Private Sub FT_DIBC_Test_400000()
    FT_DIBC_Lib_UVS256_C 400000, Array("VDD12_ADC_SOC:2.46e-05", "VDD12_AMUX_FMON:2.16e-05", "VDD12_CIO:2.45e-05", "VDD12_EFUSE1:3.60e-05", "VDD12_EFUSE2:4.00e-05", "VDD12_EFUSE3:3.40e-05", "VDD12_HSCDFT0:3.65e-05", "VDD12_HSCDFT1:3.60e-05", "VDD12_LPDP_RX:7.01e-05", "VDD12_LPDP_TX:2.37e-05", "VDD12_MIPI:2.36e-05", "VDD12_PCIE:6.76e-05", "VDD12_PCIE_REFBUF:2.21e-05", "VDD12_PLL:3.72e-05", "VDD12_PLL_DCS:2.46e-05", "VDD12_PLL_DDR:2.82e-05", "VDD12_PLL_SOC:2.21e-05", "VDD12_SLC_PLL:2.21e-05", "VDD12_ULPPLL_FLPPLL:2.31e-05", "VDD12_USB:2.22e-05", "VDD12_USB_DEBUG:2.25e-05", "VDD12_VID_PLL:2.21e-05", _
"VDD12_XTAL:2.31e-05", "VDD1:6.13e-05", "VDD2:2.38e-04", "VDDIO12_AOP:1.57e-04", "VDDIO12_AOP_2:5.66e-05", "VDDIO12_GRP:7.95e-05", "VDD_AMPH_DDR:7.99e-05", "VDD_AVE:2.27e-04", "VDD_DCS_DDR:2.15e-04", "VDD_DISP:1.81e-04", "VDD_FIXED:5.22e-05", "VDD_FIXED_AMUX:3.61e-05", "VDD_FIXED_CIO:3.40e-05", "VDD_FIXED_CPU:2.78e-05", "VDD_FIXED_ECPU_MTR:3.51e-05", "VDD_FIXED_MIPI:2.41e-05", "VDD_FIXED_MIPID_PLL:2.45e-05", "VDD_FIXED_MIPID_SOC:2.42e-05", "VDD_FIXED_MTR_CPM_PCPU:2.21e-05", "VDD_FIXED_PCIE:4.19e-05", "VDD_FIXED_PCIE_REFBUF:2.35e-05", "VDD_FIXED_PLL:2.46e-05", _
"VDD_FIXED_PLL_DCS:2.21e-05", "VDD_FIXED_PLL_SOC:2.21e-05", "VDD_FIXED_USB:2.31e-05", "VDD_FIXED_XTAL:2.21e-05", "VDD_LOW:5.85e-05", "VDD_SRAM_GPU:1.77e-04", "VDD_SRAM_ULPPLL_FLPPLL_SLC:2.51e-05", "VDD_SRAM_USB_DEBUG:2.21e-05", "VDD_SRAM_VID_PLL:2.25e-05")
End Sub

Private Sub FT_DIBC_Test_410000()
    FT_DIBC_Lib_UVS256_RELAY_UP1600 410000, Array("VDDQL_DDR", "DDR0_RREF", "K15", "DDR0_RREF", "K14", "DDR0_ZQ", "K12", "DDR0_ZQ", "K13", "DDR1_RREF", "K16", "DDR1_RREF", "K17", "DDR2_RREF", "K19", "DDR2_RREF", "K18", "DDR3_RREF", "K20", "DDR3_RREF", "K21")
End Sub

Private Sub FT_DIBC_Test_420000()
    FT_DIBC_Lib_HEXVS_C 420000, Array("VDD_CPU_SRAM:5.40e-04", "VDD_ECPU:5.46e-04", "VDD_GPU:9.89e-04", "VDD_PCPU:9.56e-04", "VDD_SOC:8.30e-04", "VDD_SRAM_SOC:5.67e-04")
End Sub

Private Sub FT_DIBC_Test_610000()
    FT_DIBC_Lib_DIFF_RELAY 610000, Array("LPDPRX1_RCAL_POS", "LPDPRX1_RCAL_NEG", "K34", "K35", "K36", "K37", "1.00e-11")
End Sub

Private Sub FT_DIBC_Test_610100()
    FT_DIBC_Lib_DIFF_RELAY 610100, Array("LPDPRX0_RCAL_POS", "LPDPRX0_RCAL_NEG", "K30", "K31", "K32", "K33", "1.00e-11")
End Sub

Private Sub FT_DIBC_Test_610200()
    FT_DIBC_Lib_DIFF_RELAY 610200, Array("ATC0_RCAL_POS", "ATC0_RCAL_NEG", "K22", "K23", "K24", "K25", "1.00e-11")
End Sub

Private Sub FT_DIBC_Test_610300()
    FT_DIBC_Lib_DIFF_RELAY 610300, Array("LPDPTX_RCAL_POS", "LPDPTX_RCAL_NEG", "K26", "K27", "K28", "K29", "1.00e-11")
End Sub

Private Sub FT_DIBC_Test_610400()
    FT_DIBC_Lib_DIFF_RELAY 610400, Array("PCIE_RCAL_P", "PCIE_RCAL_N", "K38", "K39", "K40", "K41", "1.00e-11")
End Sub

'Private Sub FT_DIBC_Test_620000()
'    FT_DIBC_Lib_FRC_BUFFER 620000, Array("NONE")
'End Sub

Private Sub FT_DIBC_Test_630000()
    FT_DIBC_Lib_LOOP_BACK 630000, Array("ATC0_RX0_N", "ATC0_TX0_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630100()
    FT_DIBC_Lib_LOOP_BACK 630100, Array("ATC0_RX0_P", "ATC0_TX0_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630200()
    FT_DIBC_Lib_LOOP_BACK 630200, Array("ATC0_RX1_N", "ATC0_TX1_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630300()
    FT_DIBC_Lib_LOOP_BACK 630300, Array("ATC0_RX1_P", "ATC0_TX1_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630400()
    FT_DIBC_Lib_LOOP_BACK 630400, Array("GP_PCIE_RX0_N", "GP_PCIE_TX0_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630500()
    FT_DIBC_Lib_LOOP_BACK 630500, Array("GP_PCIE_RX0_P", "GP_PCIE_TX0_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630600()
    FT_DIBC_Lib_LOOP_BACK 630600, Array("GP_PCIE_RX1_N", "GP_PCIE_TX1_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630700()
    FT_DIBC_Lib_LOOP_BACK 630700, Array("GP_PCIE_RX1_P", "GP_PCIE_TX1_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630800()
    FT_DIBC_Lib_LOOP_BACK 630800, Array("LPDPRX_D0_N", "LPDPRX_D1_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_630900()
    FT_DIBC_Lib_LOOP_BACK 630900, Array("LPDPRX_D0_P", "LPDPRX_D1_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631000()
    FT_DIBC_Lib_LOOP_BACK 631000, Array("LPDPRX_D2_N", "LPDPRX_D3_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631100()
    FT_DIBC_Lib_LOOP_BACK 631100, Array("LPDPRX_D2_P", "LPDPRX_D3_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631200()
    FT_DIBC_Lib_LOOP_BACK 631200, Array("LPDPRX_D4_N", "LPDPRX_D5_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631300()
    FT_DIBC_Lib_LOOP_BACK 631300, Array("LPDPRX_D4_P", "LPDPRX_D5_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631400()
    FT_DIBC_Lib_LOOP_BACK 631400, Array("LPDPRX_D6_N", "LPDPRX_D7_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631500()
    FT_DIBC_Lib_LOOP_BACK 631500, Array("LPDPRX_D6_P", "LPDPRX_D7_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631600()
    FT_DIBC_Lib_LOOP_BACK 631600, Array("LPDPRX_D8_N", "LPDPRX_D9_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631700()
    FT_DIBC_Lib_LOOP_BACK 631700, Array("LPDPRX_D8_P", "LPDPRX_D9_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631800()
    FT_DIBC_Lib_LOOP_BACK 631800, Array("ST_PCIE_RX0_N", "ST_PCIE_TX0_N", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_631900()
    FT_DIBC_Lib_LOOP_BACK 631900, Array("ST_PCIE_RX0_P", "ST_PCIE_TX0_P", "2.20e-07")
End Sub

Private Sub FT_DIBC_Test_650000()
    FT_DIBC_Lib_SB_EEPROM 650000, Array("NONE")
End Sub

Private Sub FT_DIBC_Test_660000()
    FT_DIBC_Lib_SB_Power 660000, Array()
End Sub

Private Sub FT_DIBC_Test_900000()
    FT_DIBC_Lib_Uvi80_C_Cap 900000, Array("VDD_FIXED_LPDP_RX", "C2698", "6.73e-05")
End Sub

Private Sub FT_DIBC_Test_900100()
    FT_DIBC_Lib_Uvi80_C_Cap 900100, Array("VDD_FIXED_LPDP_TX", "C2710", "3.56e-05")
End Sub

