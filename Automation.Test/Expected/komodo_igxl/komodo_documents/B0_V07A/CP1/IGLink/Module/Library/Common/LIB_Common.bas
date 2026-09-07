Attribute VB_Name = "LIB_Common"
#Const isUFP = True
Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 add bintable inital VBT.
'variable declaration
Private Const LIB_Common = "LIB_Common"
Private Const moduleName = LIB_Common
Private functionName As String

Public Const debugPrintEnable = False   'debug print in VBT modules
Public G_TestName As String 'replace testinstance for debug print CHWu 102615
Public Current_Patterns As String
Public Char_Test_Name_Curr_Loc As Long 'index for char datalog test name array
Private Type Bintable
    astrBinName() As String
    astrBinRename() As String
    astrBinSortNum() As String
End Type
Public m1_InstanceName As String

Public tyBinTable As Bintable

Public nWire_Ports_GLB As String ''Support multiple nWire port 20170718

Public Previous_DCCategory As String
Public Previous_DCSelector As String
Public DictOCR As New Scripting.Dictionary

Public Function is_reference_installed(S As String) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim x As Variant
    is_reference_installed = False
    For Each x In Application.ActiveWorkbook.VBProject.References
        If S = x.name Then
            is_reference_installed = True
            Exit Function
        End If
    Next x
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "is_reference_installed") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Function WorksheetExists(wsName As String, delete As Boolean) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim ws As Worksheet
    Dim ret As Boolean
    ret = False
    wsName = UCase(wsName)
    For Each ws In ThisWorkbook.Sheets
        If UCase(ws.name) = wsName Then
            If delete = True Then
                Application.DisplayAlerts = False
                ws.delete
                Application.DisplayAlerts = True
            End If
            ret = True
            Exit For
        End If
    Next
    WorksheetExists = ret
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "WorksheetExists") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function Wait(Time As Double, Optional debug_flag As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'pause few time
    TheHdw.Wait Time
    If Debug_Flag Then
        theexec.Datalog.WriteComment ("print: Wait time = " + CStr(Time * 1000#) + " mS")
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Wait") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



Public Function Trim_NC_Pin(ByRef original_ary() As String, ByRef original_pin_cnt As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'get active pins array
    Dim i As Long, j As Long
    Dim p As Variant
    Dim TempArray() As String
    Dim TempPinCnt As Long
    Dim NullArray() As String
    Dim TempString As String
    Dim PowerSequence As Double
    
    If original_pin_cnt <> 0 Then
        i = 0   'init
        For Each p In original_ary
            If theexec.DataManager.ChannelType(p) <> "N/C" Then i = i + 1
        Next p
        
'''''        'redim
'''''        ReDim TempArray(i - 1)
        
        j = 0   'init
        
        If i > 0 Then
            'redim
            ReDim TempArray(i - 1)
            
            For Each p In original_ary
                If theexec.DataManager.ChannelType(p) <> "N/C" Then
                    TempArray(j) = p
                    j = j + 1
                Else
                    j = j
                End If
            Next p
        End If

'''''        For Each p In original_ary
'''''            If TheExec.DataManager.ChannelType(p) <> "N/C" Then
'''''                TempArray(j) = original_ary(j)
'''''                j = j + 1
'''''            Else
'''''                j = j
'''''            End If
'''''        Next p
        
        'return array and pin count
        original_ary = TempArray
        original_pin_cnt = j
    End If
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Trim_NC_Pin") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function setXY(x As Integer, y As Integer, Optional Device As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

''''20171103 update
    Dim funcName As String:: funcName = "setXY"
    
    Dim m_chmapName As String
    Dim m_siteCnt As Long
    Dim m_match_flag As Boolean
    
    m_match_flag = False
    m_siteCnt = theexec.sites.Existing.Count
    m_chmapName = LCase(theexec.CurrentChanMap)
    
    If (m_siteCnt = 1) Then
        m_match_flag = True
        Call theexec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
        Call theexec.Datalog.Setup.WaferSetup.SetYCoord(0, y)

    ElseIf (m_siteCnt = 2) Then
        If (m_chmapName Like "*ch*2*") Then
            m_match_flag = True
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
            
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 1)
        End If
    ElseIf (m_siteCnt = 3) Then
        If (m_chmapName Like "*ch*3*") Then
            m_match_flag = True
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(1, x + 3)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(2, x + 6)
            
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(1, y)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(2, y)

            Debug.Print "Site0: " & "(" & x & "," & y & ")"
            Debug.Print "Site1: " & "(" & x + 3 & "," & y & ")"
            Debug.Print "Site2: " & "(" & x + 6 & "," & y & ")"
        End If
    ElseIf (m_siteCnt = 4) Then
        If (m_chmapName Like "*ch*4*") Then
            m_match_flag = True
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(2, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(3, x)
            
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 2)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(2, y - 4)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 6)
        End If
    ElseIf (m_siteCnt = 6) Then
        If (m_chmapName Like "*ch*6*") Then
            m_match_flag = True
             If UCase(m_chmapName) Like "*CP*" Then
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(2, x + 2)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(3, x + 2)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(4, x + 4)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(5, x + 4)
                
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 4)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 4)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 4)
                
                Debug.Print "Site0: " & "(" & x & "," & y & ")"
                Debug.Print "Site1: " & "(" & x & "," & y - 4 & ")"
                Debug.Print "Site2: " & "(" & x + 2&; "," & y & ")"
                Debug.Print "Site3: " & "(" & x + 2&; "," & y - 4 & ")"
                Debug.Print "Site4: " & "(" & x + 4&; "," & y & ")"
                Debug.Print "Site5: " & "(" & x + 4&; "," & y - 4 & ")"
            ElseIf (UCase(m_chmapName) Like "*WLFT*") Then
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(2, x + 2)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(3, x + 2)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(4, x + 4)
                Call theexec.Datalog.Setup.WaferSetup.SetXCoord(5, x + 4)
                
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 2)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 2)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
                Call theexec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 2)
                
                Debug.Print "Site0: " & "(" & x & "," & y & ")"
                Debug.Print "Site1: " & "(" & x & "," & y - 2 & ")"
                Debug.Print "Site2: " & "(" & x + 2&; "," & y & ")"
                Debug.Print "Site3: " & "(" & x + 2&; "," & y - 2&; ")"
                Debug.Print "Site4: " & "(" & x + 4&; "," & y & ")"
                Debug.Print "Site5: " & "(" & x + 4&; "," & y - 2 & ")"
            End If
            
        End If
    ElseIf (m_siteCnt = 8) Then
        If (m_chmapName Like "*ch*8*") Then
            m_match_flag = True
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(2, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(3, x)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(4, x + 2)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(5, x + 2)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(6, x + 2)
            Call theexec.Datalog.Setup.WaferSetup.SetXCoord(7, x + 2)
            
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 2)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(2, y - 4)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 6)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 2)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(6, y - 4)
            Call theexec.Datalog.Setup.WaferSetup.SetYCoord(7, y - 6)
        End If

    Else
        m_match_flag = False
    End If
    
    If (m_match_flag = False) Then
        ''''Has the reminder for user to maintain this fuction if the setup is unsuitable.
        If isDebugMode Then theexec.AddOutput "<WARNING> " + funcName + ":: The Condition Setup is Wrong."
        GoTo errHandler
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "setXY") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function RegKeySave(i_RegKey As String, i_Value As String, Optional i_Type As String = "REG_SZ")
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'add by Teradyne/Vern to output variable to RegKey
'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created
    Dim myWS As Object
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'write registry key
    i_RegKey = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\" & i_RegKey
    myWS.RegWrite i_RegKey, i_Value, i_Type
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "RegKeySave") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Function RegKeyRead(i_RegKey As String) As String

    Dim myWS As Object
    
    On Error Resume Next
    
    Set myWS = CreateObject("WScript.Shell")
    
    RegKeyRead = myWS.RegRead("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\" & i_RegKey)

End Function


Public Function Dec2Bin(ByVal n As Long, ByRef BinArray() As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim i As Integer, j As Integer
    Dim Element_Amount As Integer
    Dim Count As Integer
    '               01101
    ' BinArray(4) 1
    ' BinArray(3) 0
    ' BinArray(2) 1
    ' BinArray(1) 1
    ' BinArray(0) 0

    Element_Amount = UBound(BinArray)
    If n > (2 ^ (Element_Amount + 1) - 1) Then
        n = 0
        theexec.Datalog.WriteComment "Error(Dec2Bin): Overange for " & n
    End If

    For j = 0 To Element_Amount
        BinArray(j) = 0
    Next j

    'If n < 0 Then MsgBox ("Warning(Dec2Bin)!!! Decimal Number should be positive integer")
    If n < 0 Then
        theexec.Datalog.WriteComment " The input vlaue of (Dec2Bin) is negative, so we enforce it as 0 to prevent from error alarm."
        n = 0
    End If
    
    i = 0
    Do Until n = 0
        If (i > Element_Amount) Then theexec.Datalog.WriteComment "Warning (Dec2Bin)!!! Decimal " & n & " is over-range (>" & i & "bit)"
        If (n Mod 2) Then
            BinArray(Element_Amount - i) = 1
        Else
            BinArray(Element_Amount - i) = 0
        End If
        n = Int(n / 2)
        i = i + 1
    Loop

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Dec2Bin") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Dec2BinStr32Bit(ByVal Nbit As Long, ByVal num As Long) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    ' 2'complement: invert the number's bits and then add 1
    'Dec2BinStr32Bit 32, -65525
    '1111111111111110000000000001011    -65525
    '0000000000000001111111111110101     65525
    Dim i As Integer, j As Integer
    Dim Element_Amount As Integer
    Dim Count As Integer
    Dim binstr As String
    ' MSB "010101" LSB
    
    binstr = vbNullString
    If Nbit < 1 Then MsgBox ("Warning(Dec2BinStr32Bit)!!! Decimal Number or number of Bit is wrong")
    If Nbit = 32 Then
        Nbit = 30
        If num < 0 Then
            binstr = "1"
        Else
            binstr = "0"
        End If
    End If
    For i = Nbit To 0 Step -1
        If num And (2 ^ i) Then
            binstr = binstr & "1"
        Else
            binstr = binstr & "0"
        End If
    Next
    Dec2BinStr32Bit = binstr
'    Debug.Print BinStr
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Dec2BinStr32Bit") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function BinStr2HexStr(ByVal binstr As String, ByVal HexBit As Long) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim i As Integer, j As Integer
    Dim BinStrLen As Long
    Dim HexMOD As Integer
    Dim HexStr As String
    Dim HexVal As String
    Dim HexLen As Long

    HexStr = vbNullString
    
    BinStrLen = Len(binstr)
    If (BinStrLen Mod (4)) > 0 Then
        HexLen = (BinStrLen \ 4) + 1
    Else
        HexLen = BinStrLen \ 4
    End If
    
    If HexBit > HexLen Then
        HexLen = HexBit
    End If

    HexMOD = HexLen * 4 - BinStrLen
    
    If HexMOD > 0 Then
        For i = 0 To HexMOD - 1
            binstr = "0" & binstr
        Next i
    End If

    For i = 0 To HexLen - 1
        If mid(binstr, i * 4 + 1, 4) = "0000" Then
            HexVal = "0"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0001" Then
            HexVal = "1"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0010" Then
            HexVal = "2"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0011" Then
            HexVal = "3"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0100" Then
            HexVal = "4"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0101" Then
            HexVal = "5"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0110" Then
            HexVal = "6"
        ElseIf mid(binstr, i * 4 + 1, 4) = "0111" Then
            HexVal = "7"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1000" Then
            HexVal = "8"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1001" Then
            HexVal = "9"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1010" Then
            HexVal = "A"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1011" Then
            HexVal = "B"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1100" Then
            HexVal = "C"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1101" Then
            HexVal = "D"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1110" Then
            HexVal = "E"
        ElseIf mid(binstr, i * 4 + 1, 4) = "1111" Then
            HexVal = "F"
        Else
            HexVal = "X"
        End If

        HexStr = HexStr & HexVal
    Next i

    BinStr2HexStr = HexStr

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "BinStr2HexStr") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Function Bin2Dec(sMyBin As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec = Bin2Dec + mid(sMyBin, iLen - x + 1, 1) * 2 ^ x
    Next
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Bin2Dec") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Function Bin2Dec_rev(sMyBin As String) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec_rev = Bin2Dec_rev + mid(sMyBin, iLen - x + 1, 1) * 2 ^ (iLen - x)
    Next
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Bin2Dec_rev") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Function Bin2Dec_rev_Double(sMyBin As String) As Double
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec_rev_Double = Bin2Dec_rev_Double + mid(sMyBin, iLen - x + 1, 1) * 2 ^ (iLen - x)
    Next
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Bin2Dec_rev_Double") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function ExculdePath(Pat As Variant) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim patt_ary_temp() As String
    patt_ary_temp = Split(Pat, "\")
    ExculdePath = patt_ary_temp(UBound(patt_ary_temp))

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "ExculdePath") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''debug printing
''20230223: Modidfied to add tool check add in include by "s_OtherPrintInfo"
' [20230731][T-Har][Tank] Fix clk pin not exist, but set value error
' [20230731][All][Carter] TimeDomainIn can not set vbNullString
Public Function DebugPrintFunc(Test_Pattern As String, Optional testname_enable As Boolean = False, Optional s_OtherPrintInfo As String = vbNullString) As Long
'for debug printing generation
    Dim PinCnt As Long, PinAry() As String
    Dim i As Long
    Dim PowerVolt As Double
    Dim Powerfoldlimit As Double
    Dim AlramCheck As String
    Dim PowerAlramTime As Double
    Dim All_power_list As PinList
    Dim CurrentChans As String
    Dim PatSetArray() As String
    Dim PrintPatSet As Variant
    Dim patt As Variant 'patt1
    Dim patt1 As Variant
    Dim patt_ary_debug() As String
    Dim pat_count_debug As Long
    Dim patt_ary_debug1() As String
    Dim pat_count_debug1 As Long
    Dim PinGroup() As String
    Dim EachPinGroup As Variant
    Dim Timelist As String
    Dim TimeGroup() As String
    Dim CurrTiming As Variant
    Dim TimeDomainlist As String
    Dim TimeDomaingroup() As String
    Dim CurrTimeDomain As Variant
    Dim TimeDomainIn As String
    Dim TempString As String
    Dim TempStringOffline As String
    Dim AlarmBehavior As tlAlarmBehavior
    Dim DebugPrint_version As Double
    Dim VMain As Double
    Dim Irange As Double
    Dim Gate_State As Boolean
    Dim Gate_State_str As String
    Dim PinData As New PinListData
    Dim out_line As String
    Dim CurrSite As Variant
    Dim XI0_Vicm  As Double
    Dim XI0_Vid As Double
    Dim XI0_Vihd As Double
    Dim XI0_Vild As Double
    Dim SlotType As String
    
    Dim m_DCCategory As String
    Dim m_DCSelector As String
    Dim m_ACCategory As String
    Dim m_ACSelector As String
    Dim m_TimeSetSheet As String
    Dim m_EdgeSetSheet As String
    Dim m_LevelsSheet As String
    Dim m_tmpPMname As String
    
    Dim Pins() As String
    Dim Diff_pins() As String
    Dim pincont As Long
    Dim pincont_diff As Long
    Dim Diffenential_Pins As Variant
    Dim I_diff As Integer
    Dim J_original As Integer
    
    Dim split_powerGroup() As String ''Wyatt For Bincut pins decompose 20221027
    Dim strAry_pinSyncup() As String
    Dim cnt_DecomposedPinList As Long
    Dim PinGroup_temp As String
    Dim Bincut_pins As String
    Dim j As Integer
    
    Dim XI0_Freq_pl As New PinListData, RTCLK_Freq_pl As New PinListData, Pin_XI0 As New PinList, Pin_RTCLK As New PinList
    Dim site As Variant
    
    Dim PowerConnect_State As tlDCVSConnectWhat
    Dim PowerConnect_State_str As String
    
    Dim s_OtherPrintInfo_Ary() As String
    
    Dim sa_XI0_GP() As String
    Dim n_XI0_GP_cnt As Long
    Dim sa_XI0_Diff_GP() As String
    Dim n_XI0_Diff_GP_cnt As Long
    
    On Error GoTo errHandler
    
    'version history
    'DebugPrint_version = 1.3   'copy from Fiji
    'DebugPrint_version = 1.4   'implement offline simulation for Rhea bring up
    'DebugPrint_version = 1.5   'Update for Multi-Port nWire setting
    'DebugPrint_version = 1.6   'Add differential nWire frequency capture, DCVS tl* modes put in strings, support no pattern items
    'DebugPrint_version = 1.7   'Add DC/AC cetegory setup, remove off-limt timing simulation, offline could get real timing.
    'DebugPrint_version = 1.71   'Add PPMU debug print function.
    'DebugPrint_version = 1.72   'Add DCVI debug print support.
    DebugPrint_version = 1.73        'Add G_TestName condition to avoid empty instance name.
    Shmoo_Pattern = Test_Pattern
    m1_InstanceName = LCase(theexec.DataManager.instancename)
    
    'setups

    If DebugPrintFlag_Chk = True Then
        theexec.Datalog.WriteComment ""
        theexec.Datalog.WriteComment "================debug print start=================="
        'list all power pin's level
        theexec.Datalog.WriteComment "  DebugPrint version = " & DebugPrint_version
        If testname_enable Then
            ''20210930
            If G_TestName <> "" Then
                theexec.Datalog.WriteComment "  TestInstanceName = " & G_TestName
                testname_enable = False
            Else
                theexec.Datalog.WriteComment "  TestInstanceName = " & theexec.DataManager.instancename
            End If
        Else
            theexec.Datalog.WriteComment "  TestInstanceName = " & theexec.DataManager.instancename
        End If
        
        theexec.Datalog.WriteComment "***** List all Category info Start ******"
        ''''Get the current TestInstance Context
    
        ''''20151109
        ''''Use the local module private global variable to be flexible if it could be used anywhere in this Module. (Just in case)
        Call theexec.DataManager.GetInstanceContext(m_DCCategory, m_DCSelector, _
                                                    m_ACCategory, m_ACSelector, _
                                                    m_TimeSetSheet, m_EdgeSetSheet, _
                                                    m_LevelsSheet, vbNullString)

        ''''231213 Print EQN voltage category if the instance applys volatage from bincut.
        If s_OtherPrintInfo <> "" And s_OtherPrintInfo Like "*EQN*" Then
            m_DCCategory = Replace(s_OtherPrintInfo, " ", "_")
        End If
    
        TempString = "DC Category ="
        TempString = TempString + " " + m_DCCategory
        theexec.Datalog.WriteComment TempString
        
'        TempString = "DC Selector ="
'        TempString = TempString + " " + m_DCSelector
'        TheExec.Datalog.WriteComment TempString
        
        TempString = "AC Category ="
        TempString = TempString + " " + m_ACCategory
        theexec.Datalog.WriteComment TempString
        
'        TempString = "AC Selector ="
'        TempString = TempString + " " + m_ACSelector
'        TheExec.Datalog.WriteComment TempString
        
        TempString = "Level ="
        TempString = TempString + " " + m_LevelsSheet
        theexec.Datalog.WriteComment TempString

        TempString = "TimingSet ="
        TempString = TempString + " " + m_TimeSetSheet
        theexec.Datalog.WriteComment TempString

        theexec.Datalog.WriteComment "***** List all Category info end ******"
        theexec.Datalog.WriteComment "***** List all power Start ******"

        theexec.DataManager.DecomposePinList AllPowerPinlist, PinAry(), PinCnt
        For i = 0 To PinCnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then

                SlotType = UCase(gl_GetInstrument_Dic(LCase(PinAry(i))))
                Select Case SlotType
                    Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA, glbConstIns_VSM:
                        PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.value
                    Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                        PowerVolt = TheHdw.DCVI.Pins(PinAry(i)).Voltage
                    Case Else:
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "DebugPrintFunc", "Pin instrument is not define, so doesn't get PowerVolt!!")
                End Select
                    
                theexec.Datalog.WriteComment "  " & PinAry(i) & " = " & Format(PowerVolt, "0.000") & " v"
            End If
        Next i

        theexec.Datalog.WriteComment "***** List all power end ******"
        
        theexec.Datalog.WriteComment "***** List all Vmain power Start ******"

        For i = 0 To PinCnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then
                    SlotType = UCase(gl_GetInstrument_Dic(LCase(PinAry(i))))
                Select Case SlotType
                    Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA, glbConstIns_VSM:
                        PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Main.value
                    Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                        PowerVolt = TheHdw.DCVI.Pins(PinAry(i)).Voltage
                    Case Else:
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "DebugPrintFunc", "Pin instrument is not define, so doesn't get Vmain PowerVolt!!")
                End Select
                theexec.Datalog.WriteComment "  " & PinAry(i) & " = " & Format(PowerVolt, "0.000") & " v"
            End If
        Next i

        theexec.Datalog.WriteComment "***** List all Vmain power end ******"
        
        theexec.Datalog.WriteComment "***** List all Valt power Start ******"

        For i = 0 To PinCnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then
                SlotType = UCase(gl_GetInstrument_Dic(LCase(PinAry(i))))
                Select Case SlotType
                    Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA, glbConstIns_VSM:
                        PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Alt.value
                    Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                        PowerVolt = TheHdw.DCVI.Pins(PinAry(i)).Voltage
                    Case Else:
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "DebugPrintFunc", "Pin instrument is not define, so doesn't get Valt PowerVolt!!")
                End Select
                theexec.Datalog.WriteComment "  " & PinAry(i) & " = " & Format(PowerVolt, "0.000") & " v"
            End If
        Next i

        theexec.Datalog.WriteComment "***** List all Valt power end ******"
        
        theexec.Datalog.WriteComment "***** List all power FoldLimit TimeOut Start ******"

        TempString = "FoldLimit TimeOut :"
        For i = 0 To PinCnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then

                SlotType = UCase(gl_GetInstrument_Dic(LCase(PinAry(i))))
                Select Case SlotType
                    Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA, glbConstIns_VSM:
                        PowerAlramTime = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.TimeOut
                    Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                        PowerAlramTime = TheHdw.DCVI.Pins(PinAry(i)).FoldCurrentLimit.TimeOut
                    Case Else:
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "DebugPrintFunc", "Pin instrument is not define, so doesn't get FoldLimit TimeOut!!")
                End Select

                If i <> (PinCnt - 1) Then
                    TempString = TempString + "  " & PinAry(i) & " = " & Format(1000 * PowerAlramTime, "0.000") & " ms" + ","
                Else
                    TempString = TempString + "  " & PinAry(i) & " = " & Format(1000 * PowerAlramTime, "0.000") & " ms"
                End If
            End If
        Next i

        theexec.Datalog.WriteComment TempString
        theexec.Datalog.WriteComment "***** List all power FoldLimit TimeOut End ******"
        theexec.Datalog.WriteComment "***** List all power FoldLimit Current Start ******"

        TempString = "FoldLimit Current :"
        For i = 0 To PinCnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then

                SlotType = UCase(gl_GetInstrument_Dic(LCase(PinAry(i))))
                Select Case SlotType
                    Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA, glbConstIns_VSM:
                        Powerfoldlimit = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.level.value
                    Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                        Powerfoldlimit = TheHdw.DCVI.Pins(PinAry(i)).Current
                    Case Else:
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "DebugPrintFunc", "Pin instrument is not define, so doesn't get FoldLimit Current!!")
                End Select

                If i <> (PinCnt - 1) Then
                    TempString = TempString + "  " & PinAry(i) & " = " & Format(Powerfoldlimit, "0.000000") & " A" + ","
                Else
                    TempString = TempString + "  " & PinAry(i) & " = " & Format(Powerfoldlimit, "0.000000") & " A"
                End If
            End If
        Next i

        theexec.Datalog.WriteComment TempString
        theexec.Datalog.WriteComment "***** List all power FoldLimit Current End ******"
        theexec.Datalog.WriteComment "***** List all power Alram Check Start ******"

        TempString = "Alram Check :"
        For i = 0 To PinCnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then

                SlotType = UCase(gl_GetInstrument_Dic(LCase(PinAry(i))))
                Select Case SlotType
                    Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA, glbConstIns_VSM:
                        AlarmBehavior = TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
                    Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                        AlarmBehavior = TheHdw.DCVI.Pins(PinAry(i)).FoldCurrentLimit.Behavior
                    Case Else:
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "DebugPrintFunc", "Pin instrument is not define, so doesn't get Alram Behavior!!")
                End Select
                
                If AlarmBehavior = tlAlarmOff Then
                    AlramCheck = "tlAlarmOff"
                ElseIf AlarmBehavior = tlAlarmContinue Then
                    AlramCheck = "tlAlarmContinue"
                ElseIf AlarmBehavior = tlAlarmDefault Then
                    AlramCheck = "tlAlarmDefault"
                ElseIf AlarmBehavior = tlAlarmForceBin Then
                    AlramCheck = "tlAlarmForceBin"
                ElseIf AlarmBehavior = tlAlarmForceFail Then
                    AlramCheck = "tlAlarmForceFail"
                End If
                If i <> (PinCnt - 1) Then
                    TempString = TempString + "  " & PinAry(i) & " = " & AlramCheck & ","
                Else
                    TempString = TempString + "  " & PinAry(i) & " = " & AlramCheck
                End If
            End If
        Next i

        theexec.Datalog.WriteComment TempString
        theexec.Datalog.WriteComment "***** List all power Alram Check End ******"
        theexec.Datalog.WriteComment "***** List all power Connection Check Start ******"

        TempString = "Power Relay Connection:"
        For i = 0 To PinCnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then

                SlotType = UCase(gl_GetInstrument_Dic(LCase(PinAry(i))))
                Select Case SlotType
                    Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA, glbConstIns_VSM:
                        PowerConnect_State = TheHdw.DCVS.Pins(PinAry(i)).Connected
                    Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                        PowerConnect_State = TheHdw.DCVI.Pins(PinAry(i)).Connected
                    Case Else:
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "DebugPrintFunc", "Pin instrument is not define, so doesn't get Relay Connection!!")
                End Select
                
                Select Case PowerConnect_State
                    Case tlDCVSConnectDefault: PowerConnect_State_str = "tlDCVSConnectDefault"
                    Case tlDCVSConnectNone: PowerConnect_State_str = "tlDCVSConnectNone"
                    Case tlDCVSConnectForce: PowerConnect_State_str = "tlDCVSConnectForce"
                    Case tlDCVSConnectSense: PowerConnect_State_str = "tlDCVSConnectSense"
                End Select
                If i <> (PinCnt - 1) Then
                    TempString = TempString + "  " & PinAry(i) & " = " & PowerConnect_State_str + ","
                Else
                    TempString = TempString + "  " & PinAry(i) & " = " & PowerConnect_State_str
                End If
            End If
        Next i

        theexec.Datalog.WriteComment TempString
        theexec.Datalog.WriteComment "***** List all power Connection Check End ******"
        theexec.Datalog.WriteComment "***** List all power Gate Start ******"

        TempString = "Power Gate Status:"

        For i = 0 To PinCnt - 1
            If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then
                SlotType = UCase(gl_GetInstrument_Dic(LCase(PinAry(i))))
                Select Case SlotType
                    Case glbConstIns_HEXVS, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA, glbConstIns_VSM:
                        Gate_State = TheHdw.DCVS.Pins(PinAry(i)).Gate
                    Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                        Gate_State = TheHdw.DCVI.Pins(PinAry(i)).Gate
                    Case Else:
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "DebugPrintFunc", "Pin instrument is not define, so doesn't get Gate Status!!")
                End Select
                
                Select Case Gate_State
                    Case True: Gate_State_str = "on"
                    Case False: Gate_State_str = "off"
                End Select
                If i <> (PinCnt - 1) Then
                    TempString = TempString + "  " & PinAry(i) & " = " & Gate_State_str + ","
                Else
                    TempString = TempString + "  " & PinAry(i) & " = " & Gate_State_str
                End If
            End If
        Next i

        theexec.Datalog.WriteComment TempString
        theexec.Datalog.WriteComment "***** List all power Gate Check End ******"
        theexec.Datalog.WriteComment "***** List Pattern Start ******"

        'Print test pattern
        If Test_Pattern <> "" Then
            PatSetArray = Split(Test_Pattern, ",")

            For Each PrintPatSet In PatSetArray
                If LCase(PrintPatSet) Like "*.pat*" Then
                    theexec.Datalog.WriteComment "  Pattern : " & PrintPatSet
                Else
                    GetPatListFromPatternSet CStr(PrintPatSet), patt_ary_debug, pat_count_debug
                    For Each patt In patt_ary_debug
                        If patt <> "" Then theexec.Datalog.WriteComment "  Pattern : " & patt
                    Next patt
                End If
            Next PrintPatSet
        Else
            'do nothing, no printing
        End If

        theexec.Datalog.WriteComment "***** List Pattern end ******"
        theexec.Datalog.WriteComment "***** List Level Start ******"

        PinGroup = Split(PinGrouplist, ",")
'================================================================Mask differential pins start====================================================================
        
        Diffenential_Pins = "All_DiffPairs"
        For Each EachPinGroup In PinGroup   'EachPinGroup
            theexec.DataManager.DecomposePinList EachPinGroup, Pins, pincont
            theexec.DataManager.DecomposePinList Diffenential_Pins, Diff_pins, pincont_diff
            If DicDiffPairs.Exists(Diff_pins(0)) = False Then
                For I_diff = 0 To pincont_diff - 1
                    DicDiffPairs.Add Diff_pins(I_diff), Diff_pins(I_diff)
                Next I_diff
            End If
            For J_original = 0 To pincont - 1
                If DicDiffPairs.Exists(Pins(J_original)) = False Then
                    If gl_GetInstrument_Dic.Exists(LCase(Pins(J_original))) Then
                       theexec.Datalog.WriteComment "  Pins : " & CStr(EachPinGroup) _
                       & " , Vih = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chVih), "0.000") & " v" _
                       & " , Vil = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chVil), "0.000") & " v" _
                       & " , Voh = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chVoh), "0.000") & " v" _
                       & " , Vol = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chVol), "0.000") & " v" _
                       & " , Iol = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chIoh), "0.000") & " v" _
                       & " , Ioh = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chIol), "0.000") & " v" _
                       & " , Vt  = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chVt), "0.000") & " v" _
                       & " , Vch = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chVch), "0.000") & " v" _
                       & " , Vcl = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.value(chVcl), "0.000") & " v" _
                       & " , PPMU_VclampHi = " & Format(TheHdw.PPMU.Pins(CStr(Pins(J_original))).ClampVHi, "0.000") & " v" _
                       & " , PPMU_VclampLow = " & Format(TheHdw.PPMU.Pins(CStr(Pins(J_original))).ClampVLo, "0.000") & " v"
                       'Diff_pins_dictionary.RemoveAll
                       Exit For
                    End If
                End If
            Next J_original
'================================================================Mask differential pins end====================================================================

'            TheExec.Datalog.WriteComment "  Pins : " & CStr(EachPinGroup) _
'            & " , Vih = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVih), "0.000") & " v" _
'            & " , Vil = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVil), "0.000") & " v" _
'            & " , Voh = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVoh), "0.000") & " v" _
'            & " , Vol = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVol), "0.000") & " v" _
'            & " , Iol = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chIoh), "0.000") & " v" _
'            & " , Ioh = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chIol), "0.000") & " v" _
'            & " , Vt  = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVt), "0.000") & " v" _
'            & " , Vch = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVch), "0.000") & " v" _
'            & " , Vcl = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVcl), "0.000") & " v" _
'            & " , PPMU_VclampHi = " & Format(thehdw.PPMU.Pins(CStr(EachPinGroup)).ClampVHi, "0.000") & " v" _
'            & " , PPMU_VclampLow = " & Format(thehdw.PPMU.Pins(CStr(EachPinGroup)).ClampVLo, "0.000") & " v"
        Next EachPinGroup

        theexec.Datalog.WriteComment "***** List Level end ******"
        theexec.Datalog.WriteComment "***** List Timing Start ******"

        If Test_Pattern <> "" Then
            TimeDomainlist = TheHdw.Digital.Timing.TimeDomainlist
            TimeDomaingroup = Split(TimeDomainlist, ",")
            For Each CurrTimeDomain In TimeDomaingroup
                If CStr(CurrTimeDomain) = "All" Then
                    TimeDomainIn = ""
                Else
                    TimeDomainIn = CStr(CurrTimeDomain)
                End If
                
                Timelist = TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
                'TimeGroup
                TimeGroup = Split(Timelist, ",")
                For Each CurrTiming In TimeGroup
                    If CurrTiming <> "" Then
                        If TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.period(CStr(CurrTiming)) > 0 Then
                            theexec.Datalog.WriteComment "  Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & Format((1 / TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.period(CStr(CurrTiming))) / 1000000, "0.000") & " Mhz"
                        Else
                            theexec.Datalog.WriteComment "  Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & Format(0, "0.000") & " Mhz"
                        End If
                    End If
                Next CurrTiming
            Next CurrTimeDomain
        Else
            theexec.Datalog.WriteComment "  Time Doamin : " & "N/A" & ", TimeSet : " & "N/A" & " = " & Format(0 / 1000000, "0.000") & " Mhz"
        End If

        '' add for XI0 free running clk
'           TheExec.Datalog.WriteComment "  FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & TheHdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & TheHdw.DIB.SupportBoardClock.Vil & " v"
        theexec.DataManager.DecomposePinList XI0_GP, sa_XI0_GP, n_XI0_GP_cnt
        theexec.DataManager.DecomposePinList XI0_Diff_GP, sa_XI0_Diff_GP, n_XI0_Diff_GP_cnt
        
        If n_XI0_GP_cnt <> 0 Then 'differential(false) or single end(true)
            Pin_XI0.value = XI0_GP
            TheHdw.Digital.Pins(Pin_XI0).Levels.value(chVoh) = TheHdw.Digital.Pins(Pin_XI0).Levels.value(chVih) / 4
        ElseIf n_XI0_Diff_GP_cnt <> 0 Then
            'Vod=0, do nothing
            Pin_XI0.value = XI0_Diff_GP
            TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.value(chVod) = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.value(chVid) / 4
        End If

        If n_XI0_GP_cnt <> 0 Or n_XI0_Diff_GP_cnt <> 0 Then
            Freq_MeasFreqSetup Pin_XI0, 0.001
            Freq_MeasFreqStart Pin_XI0, 0.001, XI0_Freq_pl

            If theexec.TesterMode = testModeOffline Then
                For Each site In theexec.sites
                    XI0_Freq_pl.Pins(0).value = 24000000
                Next site
            End If
        End If

        For Each site In theexec.sites
            If n_XI0_GP_cnt <> 0 Then 'differential(false) or single end(true)
                theexec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0).Levels.value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0).Levels.value(chVil), "0.000") & " v"
                'CHWu modify 10/14 to add Xio_PA_1 and remove RTCLK
'                    TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.pins(Pin_XI0_1).Levels.Value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.pins(Pin_XI0_1).Levels.Value(chVil), "0.000") & " v"
            ElseIf n_XI0_Diff_GP_cnt <> 0 Then
              'CHWu modify 11/17 modify for Xio_PA printout
                XI0_Vicm = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.value(chVicm)
                XI0_Vid = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.value(chVid)
                XI0_Vihd = XI0_Vicm + XI0_Vid / 2
                XI0_Vild = XI0_Vicm - XI0_Vid / 2
                theexec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(XI0_Vihd, "0.000") & " v , clock_Vil: " & Format(XI0_Vild, "0.000") & " v"
'                    TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(XI0_Vihd, "0.000") & " v , clock_Vil: " & Format(XI0_Vild, "0.000") & " v"
'                    TheExec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVid), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVod), "0.000") & " v"
'                    TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0_1).DifferentialLevels.Value(chVid), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0_1).DifferentialLevels.Value(chVod), "0.000") & " v"
            End If
        Next site

        Meas_FRC vbNullString    ' Multi nWire 20170718

        theexec.Datalog.WriteComment "***** List Timing end ******"
        theexec.Datalog.WriteComment "***** List Disable Compare check Start ******"
        
        'EachPinGroup
        PinGroup = Split(PinGrouplist, ",")
        For Each EachPinGroup In PinGroup
            theexec.DataManager.DecomposePinList EachPinGroup, Pins, pincont
            For J_original = 0 To pincont - 1
                If gl_GetInstrument_Dic.Exists(LCase(Pins(J_original))) Then
                    theexec.Datalog.WriteComment "  Pins : " & CStr(Pins(J_original)) _
                    & " , Disable Compare= " & TheHdw.Digital.Pins(Pins(J_original)).DisableCompare
                    Exit For
                End If
            Next J_original
        Next EachPinGroup

        theexec.Datalog.WriteComment "***** List List Disable Compare check End ******"
        theexec.Datalog.WriteComment "***** List all utility bit status Start ******"
        theexec.DataManager.DecomposePinList All_Utility_list, PinAry(), PinCnt

        'Utility bits
        out_line = "Utility_list : "
        For Each CurrSite In theexec.sites.Active
            For i = 0 To PinCnt - 1
                If gl_GetInstrument_Dic.Exists(LCase(PinAry(i))) Then
                    PinData = TheHdw.Utility.Pins(PinAry(i)).States(tlUBStateProgrammed)    'TheHdw.Utility.pins((pinary(i)) '.States(tlUBStateCompared)
                    If i = 0 Then
                          out_line = out_line + PinAry(i) & " = " & PinData.Pins(0).value(CurrSite) '''& ","
                    Else
                          out_line = out_line & "," & PinAry(i) & " = " & PinData.Pins(0).value(CurrSite)
                    End If
                End If
            Next i
            theexec.Datalog.WriteComment out_line
            out_line = "Utility_list : "
        Next CurrSite

        theexec.Datalog.WriteComment "***** List all utility bit status end ******"
        
        
        '========================== For Bincut pins====================
        
        PinGroup_temp = vbNullString
        Bincut_pins = vbNullString
        
        If LCase(m1_InstanceName) Like "*_bv*" Or LCase(m1_InstanceName) Like "*_hbv*" Then
            split_powerGroup = Split(SyncUp_PowerPin_Group, ",")
            For i = 0 To UBound(split_powerGroup)
                Call theexec.DataManager.DecomposePinList(split_powerGroup(i), strAry_pinSyncup, cnt_DecomposedPinList)
                If cnt_DecomposedPinList > 0 Then
                    For j = 0 To cnt_DecomposedPinList - 1
                        If theexec.DataManager.NumberChannelTypesForPin(strAry_pinSyncup(j)) > 0 Then
                            '''//If the powerGroup is connected to DCVS, re-assembly the pinGroup
                            If PinGroup_temp <> "" Then
                                PinGroup_temp = PinGroup_temp & "," & strAry_pinSyncup(j)
                            Else
                                PinGroup_temp = strAry_pinSyncup(j)
                            End If
                        End If
                    Next j
                End If
            Next i
            
            theexec.Datalog.WriteComment "***** List Bincut pins Start ******"
            
            theexec.Datalog.WriteComment "Bincut pins= " & PinGroup_temp
            
            theexec.Datalog.WriteComment "***** List Bincut pins End ******"
        End If

        '''' Do nothing if s_OtherPrintInfo is EQN voltage.
        If s_OtherPrintInfo <> "" And Not s_OtherPrintInfo Like "*EQN" Then
            s_OtherPrintInfo_Ary = Split(s_OtherPrintInfo, ";")
            If UBound(s_OtherPrintInfo_Ary) >= 0 Then
                For i = 0 To UBound(s_OtherPrintInfo_Ary)
                        theexec.Datalog.WriteComment s_OtherPrintInfo_Ary(i)
                Next i
            End If
        End If
        theexec.Datalog.WriteComment "================debug print end  =================="
        theexec.Datalog.WriteComment ""
    End If

    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "DebugPrintFunc")
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function FreeRunClk_ScopeIn(PAPort As PinList, Optional DebugFlag As Boolean = False) ''update for multi nWire 20170718
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim tempStr As String
    
    Call Enable_FRC(PAPort.value, True)
    TheHdw.Wait 0.001
    
'    TheHdw.Digital.Pins(PAPort).Connect
''
''    TheHdw.Protocol.ports(PAPort).Enabled = True
''    TheHdw.Protocol.ports(PAPort).NWire.ResetPLL
''    TheHdw.Wait 0.001
''    ' Start the nWire engine.
''    Call TheHdw.Protocol.ports(PAPort).NWire.Frames("RunFreeClock").Execute
''
    If DebugFlag = True Then
        theexec.Datalog.WriteComment "print: nWire scope in, port (" & PAPort.value & ")"
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "FreeRunClk_ScopeIn") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function PowerUp_Interpose(PAPort As PinList, Optional DebugFlag As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
'    FreeRunClk_ScopeOut PAPort, DebugFlag
    'TheHdw.Utility.Pins(Relay).State = tlUtilBitOn
    
'    If DebugFlag = True Then    'debugprint
'         TheExec.Datalog.WriteComment "print: RTCLK relay on, relay " & Relay.Value
'    End If

    FreeRunClk_ScopeIn PAPort, DebugFlag
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "PowerUp_Interpose") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function PowerDown_Interpose(nWireDisconnectPin As String, Optional DebugFlag As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    FreeRunClk_Disconnect nWireDisconnectPin, DebugFlag
    'FreeRunClk_Disable nWireDisconnectPin, True 'pass site will halt also
    theexec.Datalog.WriteComment "print: nWire engine , Halt " & vbCrLf

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "PowerDown_Interpose") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Function IEDA_Initialize(ByRef Inputstr As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "IEDA_Initialize"
    
    Inputstr = vbNullString

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "IEDA_Initialize") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Function IEDA_SaveRegistry(ByVal Inputstr As String, RegistryName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "IEDA_SaveRegistry"

    Call RegKeySave(RegistryName, Inputstr)

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "IEDA_SaveRegistry") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231124][All][Tank] Check sheet exist or not
Public Function Bintable_initial()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'//=====================================================================================
    

    Dim BintableSheet As Worksheet
    Dim BinNameColumnMax As Long
    Dim BinColumnNum As Long
    Dim BinContext As String
    Dim BinColNumAccu As Long: BinColNumAccu = 0
    Dim lCount As Long
    Dim b_isHaveBintableSheet As Boolean
    b_isHaveBintableSheet = False
    Application.ScreenUpdating = False
    
    '20170529 add sheet loop to include all bin table sheets
    For Each BintableSheet In ThisWorkbook.Sheets
        If LCase(BintableSheet.name) Like "*bin*table*" Then
            #If IGXL8p30 Then
            #Else
                BintableSheet.Activate
            #End If
            BinNameColumnMax = BintableSheet.Cells(Rows.Count, 2).End(xlUp).Row
            BinNameColumnMax = Worksheets(BintableSheet.name).UsedRange.Rows.Count
            ReDim Preserve tyBinTable.astrBinName(BinNameColumnMax - 4 + BinColNumAccu)
            ReDim Preserve tyBinTable.astrBinRename(BinNameColumnMax - 4 + BinColNumAccu)
            ReDim Preserve tyBinTable.astrBinSortNum(BinNameColumnMax - 4 + BinColNumAccu)
            
            For BinColumnNum = 4 To BinNameColumnMax
                BinContext = BintableSheet.Cells(BinColumnNum, 2).value
                If BinContext <> "" Then
                    tyBinTable.astrBinName(BinColumnNum - 4 + BinColNumAccu) = BinContext
                    tyBinTable.astrBinRename(BinColumnNum - 4 + BinColNumAccu) = BintableSheet.Cells(BinColumnNum, 1).value
                    tyBinTable.astrBinSortNum(BinColumnNum - 4 + BinColNumAccu) = BintableSheet.Cells(BinColumnNum, 5).value
                Else
                    Exit For
                End If
            Next BinColumnNum
        
            BinColNumAccu = BinColumnNum - 4 + BinColNumAccu
            b_isHaveBintableSheet = True
        End If
    Next BintableSheet
    Application.ScreenUpdating = True
    
    If b_isHaveBintableSheet Then
            
        For lCount = 0 To UBound(tyBinTable.astrBinName)
            If tyBinTable.astrBinSortNum(lCount) <> "" Then
                If tyBinTable.astrBinRename(lCount) <> "" Then
                    Call theexec.Datalog.SBRFill(tyBinTable.astrBinSortNum(lCount), tyBinTable.astrBinRename(lCount))
                Else
                    Call theexec.Datalog.SBRFill(tyBinTable.astrBinSortNum(lCount), mid(tyBinTable.astrBinName(lCount), InStr(tyBinTable.astrBinName(lCount), "_") + 1))
                End If
            End If
        Next lCount
    Else
        theexec.Datalog.WriteComment "There is no 'Bintable' Sheet in this workbook"
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Bintable_initial") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function DebugPrintFunc_PPMU(PPMU_Pins As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'for debug printing generation
    Dim PinCnt As Long, PinAry() As String
    Dim i As Long
    Dim PowerVolt As Double
    Dim PowerCurrent As Double
    Dim Powerfoldlimit As Double
    Dim AlramCheck As String
    Dim PowerAlramTime As Double
    Dim All_power_list As PinList
    Dim PinGroup() As String
    Dim EachPinGroup As Variant
    Dim DebugPrint_version As Double
    Dim Pins() As String, Pin_Cnt As Long
    Dim PPMU_used_Pin As Variant
    Dim PPMU_ForceV As String
    Dim PPMU_Forcei As String
    Dim DCVI_Mode As String
    Dim DCVI_sense_relay As Boolean
    Dim DCVI_force_relay As Boolean
    
    
    'version history
    DebugPrint_version = 1.71   'implement PPMU debug print function
    
    'setups

    If DebugPrintFlag_Chk = True Then
        
        If PPMU_Pins <> "" Then
        
            theexec.DataManager.DecomposePinList PPMU_Pins, Pins(), Pin_Cnt
        Else
            Pin_Cnt = 0
        End If
        
        theexec.Datalog.WriteComment ""
        theexec.Datalog.WriteComment "================debug print PPMU start=================="
        'list all power pin's level
        theexec.Datalog.WriteComment "  DebugPrint version = " & DebugPrint_version
        theexec.Datalog.WriteComment "  TestInstanceName = " & theexec.DataManager.instancename
        theexec.Datalog.WriteComment "***** List all power Start ******"
        theexec.DataManager.DecomposePinList AllPowerPinlist, PinAry(), PinCnt
        For i = 0 To PinCnt - 1
            If theexec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then

                PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Main.value
                    
                theexec.Datalog.WriteComment "  " & PinAry(i) & " = " & Format(PowerVolt, "0.000") & " v"
            End If
        Next i
        theexec.Datalog.WriteComment "***** List all power end ******"


        If glb_TesterType = "Jaguar" Then
            theexec.Datalog.WriteComment "***** List all DCVI Start ******"
            theexec.DataManager.DecomposePinList AllDCVIPinlist, PinAry(), PinCnt
            For i = 0 To PinCnt - 1
                If theexec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
        
                    PowerVolt = TheHdw.DCVI.Pins(PinAry(i)).Voltage
                    PowerCurrent = TheHdw.DCVI.Pins(PinAry(i)).Current
                    
                    If TheHdw.DCVI.Pins(PinAry(i)).mode = tlDCVIModeVoltage Then
                        DCVI_Mode = "ForceV"
                    ElseIf TheHdw.DCVI.Pins(PinAry(i)).mode = tlDCVIModeCurrent Then
                        DCVI_Mode = "ForceI"
                    Else
                        DCVI_Mode = "HighImpedance"
                    End If
        
        
                    If TheHdw.DCVI.Pins(PinAry(i)).Connected = 0 Then
                        DCVI_force_relay = False
                        DCVI_sense_relay = False
                    ElseIf TheHdw.DCVI.Pins(PinAry(i)).Connected = 1 Then
                        DCVI_force_relay = True
                        DCVI_sense_relay = False
                    ElseIf TheHdw.DCVI.Pins(PinAry(i)).Connected = 2 Then
                        DCVI_force_relay = False
                        DCVI_sense_relay = True
                    ElseIf TheHdw.DCVI.Pins(PinAry(i)).Connected = 3 Then
                        DCVI_force_relay = True
                        DCVI_sense_relay = True
                    End If
        
                    theexec.Datalog.WriteComment "  DCVI_Pins : " & PinAry(i) _
                    & " , Voltage = " & Format(PowerVolt, "0.000000") & " v" _
                    & " , Current = " & Format(PowerCurrent, "0.000000") & " A" _
                    & " , Mode = " & DCVI_Mode & " " _
                    & " , Gate = " & TheHdw.DCVI.Pins(PinAry(i)).Gate _
                    & " , DCVI_sense_relay = " & DCVI_sense_relay _
                    & " , DCVI_force_relay = " & DCVI_force_relay
                    
                End If
            Next i
            theexec.Datalog.WriteComment "***** List all DCVI end ******"
        End If
        
        theexec.Datalog.WriteComment "***** List PPMU condition Start ******"

        If Pin_Cnt > 0 Then
            For Each PPMU_used_Pin In Pins
                If theexec.DataManager.ChannelType(PPMU_used_Pin) <> "N/C" Then
                    PPMU_ForceV = CStr(Format(TheHdw.PPMU.Pins(PPMU_used_Pin).Voltage.value, "0.000000"))
                    PPMU_Forcei = CStr(Format(TheHdw.PPMU.Pins(PPMU_used_Pin).Current.value, "0.000000"))
                    
                    If TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).mode = tlPPMUForceVMeasureI Then
                        PPMU_Forcei = "None"
                    Else
                        PPMU_ForceV = "None"
                    End If

                    theexec.Datalog.WriteComment "  Pins : " & CStr(PPMU_used_Pin) _
                    & " , PPMU_VclampHi = " & Format(TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).ClampVHi, "0.000") & " v" _
                    & " , PPMU_VclampLow = " & Format(TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).ClampVLo, "0.000") & " v" _
                    & " , PPMU_forceV = " & PPMU_ForceV & " v" _
                    & " , PPMU_ForceI = " & PPMU_Forcei & " A"
              End If
            Next PPMU_used_Pin
        End If
        
        theexec.Datalog.WriteComment "***** List PPMU condition end ******"
            

        theexec.Datalog.WriteComment "================debug print PPMU end  =================="
        theexec.Datalog.WriteComment ""
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "DebugPrintFunc_PPMU") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function GetPatFromPatternSet(TestPat As String, _
                              rtnPatNames() As String, _
                              rtnPatCnt As Long) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim PatCnt As Long                          '<- Number of patterns in set
    Dim RawNameData() As String                 '<- Raw pattern name data
    Dim rtnPatNames1() As String
    Dim rtnPatNames2() As String
    Dim i As Long, j As Long
    '___ Init _____________________________________________________________________________
'    On Error GoTo errhandler
    
    '___ Check the name ___________________________________________________________________
    '    Individual pattern name or non-pattern string returns an error - thus false
    '--------------------------------------------------------------------------------------
    rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
    If (UBound(rtnPatNames) > 0) Then
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            GetPatFromPatternSet = True
            rtnPatCnt = UBound(rtnPatNames) + 1
        Else
            rtnPatCnt = 0
            For i = 0 To UBound(rtnPatNames)
                rtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames(i), PatCnt)
                rtnPatCnt = rtnPatCnt + UBound(rtnPatNames2) + 1
            Next i
            rtnPatNames1 = TheExec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
            ReDim rtnPatNames(rtnPatCnt)
            rtnPatCnt = 0
            For i = 0 To UBound(rtnPatNames1)
                rtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames1(i), PatCnt)
                For j = 0 To UBound(rtnPatNames2)
                    If LCase(rtnPatNames2(j)) Like "*.pat*" Then
                        rtnPatNames(rtnPatCnt) = rtnPatNames2(j)
                    Else
                        TheExec.ErrorLogMessage TestPat & " in more than 2 level of pattern set"
                    End If
                    rtnPatCnt = rtnPatCnt + 1
                Next j
            Next i
            GetPatFromPatternSet = True
        End If
    Else
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            GetPatFromPatternSet = True
            rtnPatCnt = 1
        Else
            rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames(0), PatCnt)
            rtnPatCnt = UBound(rtnPatNames) + 1
            For j = 0 To UBound(rtnPatNames)
                If LCase(rtnPatNames(j)) Like "*.pat*" Then
                Else
                    TheExec.ErrorLogMessage TestPat & " in more than 2 level of pattern set"
                End If
            Next j
        End If
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    GetPatFromPatternSet = False
    rtnPatCnt = -1
    Call Print_Error_Message(Error_Info, "LIB_Common", "GetPatFromPatternSet") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function FreeRunClk_Disconnect(nWireDisconnectPin As String, Optional DebugFlag As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "FreeRunClk_Disconnect"

    TheHdw.Digital.Pins(nWireDisconnectPin).Disconnect
    
    If DebugFlag = True Then theexec.Datalog.WriteComment "print: nWire disconnect, pin " & nWireDisconnectPin

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "FreeRunClk_Disconnect") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Find_nWire_Pin() As Long  ''Support multiple nWire port 20170718
' Get all nWire port and put in global variable nWire_Ports_GLB
On Error GoTo errHandler

    Dim sFuncName As String:: sFuncName = "Find_nWire_Pin"
    Dim i As Long
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim row_cnt As Long
    Dim nWire_cnt As Long
    Dim nWire_Pin_ary(10) As String
    Dim curr_pin As String, last_pin As String
        
    If gl_isFind_nWire_Pin = False Then

        Application.ScreenUpdating = False
    '    If nWire_Ports_GLB <> "" Then Exit Function
        nWire_Ports_GLB = vbNullString
        
        Set wb = Application.ActiveWorkbook
        Set ws = wb.Sheets("Levels_nWire")
        ws.Activate

    '    row_cnt = ws.Cells(Rows.Count, 2).End(xlUp).Row - 1
        nWire_cnt = 0
        
        '20210416,Add forUFp
        If glb_TesterType = "Jaguar" Then
            For i = 4 To Rows.Count 'skip header line
                curr_pin = ws.Cells(i, 2)
                If curr_pin = "" Then
                    i = Rows.Count + 1 'stop at empty row/cell
                ElseIf curr_pin Like "*_PA" And curr_pin <> last_pin Then
                    nWire_Pin_ary(nWire_cnt) = curr_pin
                    last_pin = curr_pin
                    nWire_cnt = nWire_cnt + 1
                End If
            Next i
                
        ElseIf glb_TesterType = "UltraFLEXplus" Then
            For i = 4 To Rows.Count 'skip header line
                curr_pin = ws.Cells(i, 2)
                If curr_pin = "" Then
                    i = Rows.Count + 1 'stop at empty row/cell
                ElseIf UCase(curr_pin) Like "*_PA" Or UCase(curr_pin) Like "REFCLK_*" Then
                    i = Rows.Count + 1
                ElseIf curr_pin <> last_pin Then
                    nWire_Pin_ary(nWire_cnt) = curr_pin
                    last_pin = curr_pin
                    nWire_cnt = nWire_cnt + 1
                End If
            Next i
                
        End If
        
        For i = 1 To nWire_cnt
            If nWire_Ports_GLB <> "" Then
                nWire_Ports_GLB = nWire_Ports_GLB & "," & nWire_Pin_ary(i - 1)
            Else
                nWire_Ports_GLB = nWire_Pin_ary(i - 1)
            End If
        Next i
        gl_isFind_nWire_Pin = True
        Application.ScreenUpdating = True
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, sFuncName)
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20231124][All][Tank] Fix UFP get port name error
Public Function Get_nWire_Name(NWire As Variant, port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_PowerSequence_pa As String) ''Support multiple nWire port 20170718
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
' Nwire is input and can be port name or pin name
' Will output relative name for port/ac_spec/pin/global_spec for powerup_seq
    'Eg. nWire = "XI0_Port"
    'XI0_Port,XI0_Freq_VAR, XI0_PA
    'XI0_Diff_Port,XI0_Diff_Freq_VAR, XI0_Diff_PA
    Dim remove_name As String, key_name As String
    NWire = LCase(NWire)
    If NWire Like "*_port" Then
        remove_name = "_port"
    ElseIf NWire Like "*_freq_var" Then
        remove_name = "_freq_var"
    ElseIf NWire Like "*_pa" Then
        remove_name = "_pa"
    Else
'        TheExec.ErrorLogMessage NWire & "is Wrong nWire name (should be as A_port, A_Freq_Var or A_PA)"
        remove_name = vbNullString 'if it is XI0/RT_CLK32768/CLK_IN ...
    End If
    key_name = UCase(Replace(NWire, remove_name, ""))
    If glb_TesterType = "Jaguar" Then
        port_pa = key_name & "_PORT"
        ac_spec_pa = key_name & "_FREQ_VAR"
        global_spec_PowerSequence_pa = key_name & "_Port_PowerSequence_GLB"
        pin_pa = key_name & "_PA"
        
    ElseIf glb_TesterType = "UltraFLEXplus" Then
        NWire = LCase(NWire)
        port_pa = NWire         ''port_pa = key_name & "_PORT" '' No Need _PORT different pin name in UFP
        ac_spec_pa = key_name & "_FREQ_VAR"
        global_spec_PowerSequence_pa = NWire & "_PowerSequence_GLB"
        pin_pa = NWire    ''pin_pa = Replace(key_name, "_DIFF", "_PA") '' No Need _PA different pin name in UFP
        
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Get_nWire_Name") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Disable_FRC(nWire_ports As String, Optional DisConnectFRC As Boolean = False) ''Support multiple nWire port 20170718
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
' nWire_ports  can be port name or pin name
' If it is blank, will assume to use all nWire ports
    'Eg. nWire_ports = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    
    Dim funcName As String:: funcName = "Disable_FRC"
    
    Dim nWire_port_ary() As String
    Dim nwp As Variant, all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim site As Variant
    If nWire_ports = "" Then nWire_ports = nWire_Ports_GLB
    nWire_port_ary = Split(nWire_ports, ",")
    ' Convert nWire_ports to all_ports and all_pins
    For Each nwp In nWire_port_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        If all_ports = "" Then
            all_ports = port_pa
            all_pins = pin_pa
        Else
            all_ports = all_ports & "," & port_pa
            all_pins = all_pins & "," & pin_pa
        End If
    Next nwp
    If glb_TesterType = "Jaguar" Then ''' Add for UFP
        theexec.Datalog.WriteComment "******************  Disable freerunning clock " & all_ports & " ****************"
        For Each site In theexec.sites
            TheHdw.Protocol.ports(all_ports).Halt
            TheHdw.Protocol.ports(all_ports).Enabled = False
        Next site
        
        If DisConnectFRC = True Then
            theexec.Datalog.WriteComment "******************  Disconnect nWire pins " & all_pins & " ****************"
            TheHdw.Digital.Pins(all_pins).Disconnect
        End If
    ElseIf glb_TesterType = "UltraFLEXplus" Then
        theexec.Datalog.WriteComment "******************  Disable freerunning clock " & all_ports & " ****************"
        If TheHdw.Digital.Pins(all_pins).FreeRunningClock.IsRunning Then
            TheHdw.Digital.Pins(all_pins).FreeRunningClock.stop
            TheHdw.Digital.Pins(all_pins).FreeRunningClock.Enabled = False
        End If
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Disable_FRC") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''20230223: Modidfied to change FRC frequency by instance in PLUS.
''[20230406][T-Ibi] Add gl_nWireFreq_Value_Dict to that CZ use
''[20230809][ALL] Add Freq & Threshold_Range argument to Function Enable_FRC
' [20230906][T-Spa][Jim] SBC Free Running Clock gating
Public Function Enable_FRC(nWires As String, Optional ConnectFRC As Boolean = False, Optional Freq As Double, Optional Threshold_Range As String) ''Support multiple nWire port 20170718
' nWires  can be port name or pin name
' If it is blank, will assume to use all nWire ports
    'Eg. nWires = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    On Error GoTo errHandler
    
    Dim sFuncName As String:: sFuncName = "Enable_FRC"
    
    Dim nWires_ary() As String
    Dim nwp As Variant, all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim PLL_Lock As New SiteLong
    Dim port_level_value As Double
    Dim FreeRunFreq As Double
    Dim Flag_IsPLLLocked As Boolean
    Dim site As Variant

    If nWires = "" Then nWires = nWire_Ports_GLB
    nWires_ary = Split(nWires, ",")
    ' Convert nWires to all_ports and all_pins
    
    For Each nwp In nWires_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        If all_ports = "" Then
            all_ports = port_pa
            all_pins = pin_pa
        Else
            all_ports = all_ports & "," & port_pa
            all_pins = all_pins & "," & pin_pa
        End If
    Next nwp
    
    If ConnectFRC = True Then
        TheHdw.Digital.Pins(all_pins).Connect
        theexec.Datalog.WriteComment "Connect nWire pins " & all_pins
    End If
    
    nWires_ary = Split(all_ports, ",")
    
    '''-----------------UF-----------------
    If UCase(glb_TesterType) = UCase("Jaguar") Then
        TheHdw.Protocol.ports(all_ports).Enabled = True

        For Each nwp In nWires_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            If TheHdw.Protocol.ports(nwp).Family = "FRC" Then
                TheHdw.Protocol.ports(nwp).FRC.ResetPLL
                TheHdw.Wait 0.001
                If TheHdw.Protocol.ports(nwp).Status = tlProtocolPortStatus_Running Then
                Else
                    Call TheHdw.Protocol.ports(nwp).FRC.start
                End If
            Else '' nWire
                TheHdw.Protocol.ports(nwp).NWire.ResetPLL
                TheHdw.Wait 0.001
                Call TheHdw.Protocol.ports(nwp).NWire.Frames("RunFreeClock").Execute
                TheHdw.Protocol.ports(nwp).IdleWait
            End If
        Next nwp
        If (glb_FlagCheckingFRCClock = False) And (theexec.TesterMode = testModeOnline) Then
           Call FRC_Compare(nWires, Freq, Threshold_Range)
        End If
    End If
    '''-----------------UF-----------------

    theexec.Datalog.WriteComment "Enable nWire Clock " & all_ports
    '****print out to data log about nWire clock condition
    
    For Each nwp In nWires_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        
        '''-----------------UF-----------------
        If UCase(glb_TesterType) = UCase("Jaguar") Then
            For Each site In theexec.sites
    
                If TheHdw.Protocol.ports(nwp).Family = "FRC" Then
                    Flag_IsPLLLocked = TheHdw.Protocol.ports(nwp).FRC.IsPLLLocked
                Else '' nWire
                    Flag_IsPLLLocked = TheHdw.Protocol.ports(nwp).NWire.IsPLLLocked
                End If
                
                If Flag_IsPLLLocked = False Then
                    PLL_Lock = 0
                Else
                    PLL_Lock = 1
                End If

            Next site
            
            FreeRunFreq = 1 / TheHdw.Digital.Timing.period(nwp) / 1000000
            If theexec.TesterMode = testModeOffline Then
                For Each site In theexec.sites.Selected
                    PLL_Lock = 1
                    FreeRunFreq = theexec.Specs.AC.item(ac_spec_pa).CurrentValue / 1000000  'offline
                 Next site
            End If
        
            
            theexec.Flow.TestLimit PLL_Lock, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="nWire " & nwp & " PLL_Lock" 'BurstResult=1:Pass
            If LCase(nwp) Like "*diff*" Then
                port_level_value = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.value(chVid)
                theexec.Datalog.WriteComment "********** freerunning clock(" & nwp & ") = " & Format(FreeRunFreq, "0.000") & " Mhz, Vid = " & port_level_value
            Else
                port_level_value = TheHdw.Digital.Pins(pin_pa).Levels.value(chVih)
                theexec.Datalog.WriteComment "********** freerunning clock(" & nwp & ") = " & Format(FreeRunFreq, "0.000") & " Mhz, Vih = " & port_level_value
            End If
        
        ElseIf glb_TesterType = "UltraFLEXplus" Then
        
            If TheHdw.Digital.Pins(pin_pa).FreeRunningClock.IsRunning Then 'ADR added check for clock running already
                ' stop it before setting up
                TheHdw.Digital.Pins(pin_pa).FreeRunningClock.stop
            End If
            
            With TheHdw.Digital.Pins(pin_pa)
                ac_spec_pa = Replace(ac_spec_pa, "_VAR", "_GLB")
                ' Add gl_nWireFreq to pass Frequency from Argument 221013
                If gl_nWireFreq = 0 Then gl_nWireFreq = theexec.Specs.Globals.item(ac_spec_pa).ContextValue
                .FreeRunningClock.Frequency = gl_nWireFreq
                FreeRunFreq = .FreeRunningClock.Frequency
                .FreeRunningClock.Enabled = True
                .Connect
                .FreeRunningClock.start
            End With
             
        End If
        '''-----------------UFP----------------
        If gl_nWireFreq_Value_Dict.Exists(pin_pa) Then
            gl_nWireFreq_Value_Dict.Remove pin_pa
            gl_nWireFreq_AC_Dict.Remove pin_pa
        End If
        gl_nWireFreq_Value_Dict.Add pin_pa, gl_nWireFreq ''Add for CZ PrintShmooInfo, 20230316
        gl_nWireFreq_AC_Dict.Add pin_pa, ac_spec_pa ''Add for CZ PrintShmooInfo, 20230316
        gl_nWireFreq = 0
    Next nwp
    
    '''-----------------UFP-----------------
    If UCase(glb_TesterType) = UCase("UltraFLEXplus") Then
        theexec.Datalog.WriteComment "clk pin : " & pin_pa & " ,frequency : " & FreeRunFreq & "  Start" ''Print Freq in DataLog
    End If
    '''-----------------UFP-----------------
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, sFuncName)      '20221205 Tank try use print error
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function Meas_FRC(nWire_ports As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'Eg. nWire_ports = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    Dim nWire_port_ary() As String
    Dim nwp As Variant, meas_freq As New PinListData, site As Variant
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim PA_Vicm As Double, PA_Vid As Double, PA_Vihd As Double, PA_Vild As Double
    Dim pinlist_pa As New PinList
    
    If nWire_ports = "" Then nWire_ports = nWire_Ports_GLB
    nWire_port_ary = Split(nWire_ports, ",")
    For Each nwp In nWire_port_ary
        Get_nWire_Name CStr(nwp), port_pa, ac_spec_pa, pin_pa, global_spec_pa
        If port_pa Like "*DIFF*" Then
            TheHdw.Digital.Pins(pin_pa).DifferentialLevels.value(chVod) = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.value(chVid) / 4
        Else
            TheHdw.Digital.Pins(pin_pa).Levels.value(chVoh) = TheHdw.Digital.Pins(pin_pa).Levels.value(chVih) / 4
        End If
        pinlist_pa.value = pin_pa
        Freq_MeasFreqSetup pinlist_pa, 0.001
        Freq_MeasFreqStart pinlist_pa, 0.001, meas_freq
            
        If theexec.TesterMode = testModeOffline Then
            For Each site In theexec.sites
                meas_freq.Pins(0).value = theexec.Specs.AC(ac_spec_pa).CurrentValue
            Next site
        End If
                        
        For Each site In theexec.sites
            If port_pa Like "*DIFF*" Then
                PA_Vicm = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.value(chVicm)
                PA_Vid = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.value(chVid)
                PA_Vihd = PA_Vicm + PA_Vid / 2
                PA_Vild = PA_Vicm - PA_Vid / 2
                theexec.Datalog.WriteComment "  FreeRunFreq (" & pin_pa & ") : " & Format(meas_freq.Pins(0).value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(PA_Vihd, "0.000") & " v , clock_Vil: " & Format(PA_Vild, "0.000") & " v"
            Else
                theexec.Datalog.WriteComment "  FreeRunFreq (" & pin_pa & ") : " & Format(meas_freq.Pins(pin_pa).value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(pin_pa).Levels.value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(pin_pa).Levels.value(chVil), "0.000") & " v"
            End If

        Next site
    Next nwp
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Meas_FRC") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


'[20230511][T-All][Tank] add Exit function
'[20230809][T-All][Oliver] remove the *.Reinitialize for DCVS and DCVI instrument setting
Public Function StartProfile(PinName As String, WhatToCapture As String, SampleRate As Double, sampleSize As Double, CapSignalName As String, Slot_Type As String, _
                                Optional Meter_I_Range As Double, Optional Meter_V_Range As Double)
On Error GoTo errHandler
    Dim i As Long
    Dim PinCnt As Long
    Dim PinAry() As String
    
    ' Wait if another capture is running
    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop
    
    ' Clear capture memory
    TheHdw.DCVS.Pins(PinName).ClearCaptureMemory
    
    ' Create a SIGNAL to set up instrument
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Add CapSignalName

    ' Set this as the default signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.DefaultSignal = CapSignalName
    
    theexec.DataManager.DecomposePinList PinName, PinAry(), PinCnt
    
    ' Define the signal used for the capture
    For i = 0 To PinCnt - 1
        With TheHdw.DCVS.Pins(PinAry(i)).Capture.Signals.item(CapSignalName)
'            .Reinitialize
            If (UCase(WhatToCapture) = "I") Then
                .mode = tlDCVSMeterCurrent
                .range = TheHdw.DCVS.Pins(PinAry(i)).Meter.CurrentRange.max
            Else
                .mode = tlDCVSMeterVoltage
                .range = TheHdw.DCVS.Pins(PinAry(i)).Meter.VoltageRange.max
            End If
            
            .SampleRate = SampleRate
            .sampleSize = sampleSize
    
        End With
    Next i
    ' Setup the hardware by loading the signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.item(CapSignalName).LoadSettings

    ' Start the capture
    TheHdw.DCVS.Pins(PinName).Capture.Signals.item(CapSignalName).Trigger
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "StartProfile")
    If AbortTest Then Exit Function Else Resume Next
End Function


'[20230406][RF] Add DCVI plot current profile
'[20230809][T-All][Oliver] remove the *.Reinitialize for DCVS and DCVI instrument setting
Public Function StartProfile_DCVI(PinName As String, WhatToCapture As String, SampleRate As Double, sampleSize As Double, CapSignalName As String, Slot_Type As String, _
                                    Optional Meter_I_Range As Double, Optional Meter_V_Range As Double)
On Error GoTo errHandler
    Dim i As Long
    Dim PinCnt As Long
    Dim PinAry() As String
    Dim p As Variant
    
    ' Wait if another capture is running
    Do While TheHdw.DCVI.Pins(PinName).Capture.IsCaptureDone = False
    Loop
    
    ' Clear capture memory
'    TheHdw.DCVI.Pins(PinName).Capture.Signals.Reinitialize
    
    ' Create a SIGNAL to set up instrument
    TheHdw.DCVI.Pins(PinName).Capture.Signals.Add CapSignalName

    ' Set this as the default signal
    TheHdw.DCVI.Pins(PinName).Capture.Signals.DefaultSignal = CapSignalName
    
    theexec.DataManager.DecomposePinList PinName, PinAry(), PinCnt
    
    ' Define the signal used for the capture
    For Each p In PinAry
        With TheHdw.DCVI.Pins(p).Capture.Signals.item(CapSignalName)
'            .Reinitialize
            If (UCase(WhatToCapture) = "I") Then
                .mode = tlDCVIMeterCurrent
                .range = TheHdw.DCVI.Pins(p).Meter.CurrentRange.max
            Else
                .mode = tlDCVIMeterVoltage
                .range = TheHdw.DCVI.Pins(p).Meter.VoltageRange.max
            End If
            
            .SampleRate = SampleRate
            .sampleSize = sampleSize

        End With
    Next p

    ' Setup the hardware by loading the signal
    TheHdw.DCVI.Pins(PinName).Capture.Signals.item(CapSignalName).LoadSettings

    ' Start the capture
    TheHdw.DCVI.Pins(PinName).Capture.Signals.item(CapSignalName).Trigger

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "StartProfile_DCVI")
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20230420][All][Tank] modify instrument name
Public Function SplitPinByinstrument(PinName As String, ByRef HexPins As String, ByRef UVSPins As String, ByRef VSMPins As String, ByRef VS5APins As String, ByRef VS800mAPins As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim Profile_AllPin() As String
    Dim PinCnt As Long
    Dim pin As Variant
    Dim SlotType As String
    Dim s_Msg As String

    theexec.DataManager.DecomposePinList PinName, Profile_AllPin(), PinCnt
        
    For Each pin In Profile_AllPin
        SlotType = UCase(GetInstrument(CStr(pin), 0))
        If (SlotType = glbConstIns_HEXVS) Then
        
            If HexPins = "" Then
                HexPins = pin
            Else
                HexPins = HexPins + "," + pin
            End If
        ElseIf (SlotType = glbConstIns_VHDVS) Then
            
            If UVSPins = "" Then
                UVSPins = pin
            Else
                UVSPins = UVSPins + "," + pin
            End If
        ElseIf (SlotType = glbConstIns_VSM) Then
            
            If VSMPins = "" Then
                VSMPins = pin
            Else
                VSMPins = VSMPins + "," + pin
            End If
        ElseIf (SlotType = glbConstIns_VS5A) Then
            
            If VS5APins = "" Then
                VS5APins = pin
            Else
                VS5APins = VS5APins + "," + pin
            End If
        ElseIf (SlotType = glbConstIns_VS800MA) Then
            
            If VS800mAPins = "" Then
                VS800mAPins = pin
            Else
                VS800mAPins = VS800mAPins + "," + pin
            End If
            
        Else
            s_Msg = "Pin = " & pin & " not define in function!!"
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "SplitPinByinstrument", s_Msg)
        End If
    Next pin

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "SplitPinByinstrument") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20230420][All][Tank] modify instrument name
' [20230809][All][Oliver] modify for current profile other sheet method
Public Function ProfileAutoResolution(SlotType As String, measuretime As Double, ByRef sampleSize As Double, ByRef SampleRate As Double, Optional DownSampleRatio As Long = 1, Optional nSpecificSampleRate As Long = 0)
    On Error GoTo errHandler
    Dim SampleSize_Ratio As Long: SampleSize_Ratio = 0
    '**************************HexVs***********************************
    Dim HexVsMaxSampleSize As Double: HexVsMaxSampleSize = 256000#
    Dim HexVsMaxSampleRate As Double: HexVsMaxSampleRate = 25000000#
    Dim HexMaxtime As Double: HexMaxtime = HexVsMaxSampleSize * 1024 / HexVsMaxSampleRate
    '******************************************************************
    
    '****************************UVS***********************************
    Dim UVSMaxSampleSize As Double: UVSMaxSampleSize = 16384
    Dim UVSMaxSampleRate As Double: UVSMaxSampleRate = 200000#
    Dim UVSMaxtime As Double: UVSMaxtime = UVSMaxSampleSize * (2 ^ 12) / UVSMaxSampleRate
    '******************************************************************
    
    '****************************VSM***********************************
    Dim VSMMaxSampleSize As Double: VSMMaxSampleSize = 256000#
    Dim VSMMaxSampleRate As Double: VSMMaxSampleRate = 31250000#
    Dim VSMMaxtime As Double: VSMMaxtime = VSMMaxSampleSize * (2 ^ 12) / VSMMaxSampleRate
    '******************************************************************
    
    '****************************VS-5A/VS-800mA***********************************
    Dim VSMaxSampleSize As Double: VSMaxSampleSize = 262143
    Dim VSMaxSampleRate As Double: VSMaxSampleRate = 200000#
    Dim VSMaxtime As Double: VSMaxtime = VSMaxSampleSize * (2 ^ 12) / VSMaxSampleRate
    '******************************************************************
    
    
    Dim RealRate As Double
    Dim i As Integer
    Dim prediff As Double
    Dim posdiff As Double
''Time * SampleRate = SampleSize
    Select Case SlotType
        Case glbConstIns_HEXVS
            If measuretime > HexMaxtime Then
                sampleSize = HexVsMaxSampleSize
                SampleRate = HexVsMaxSampleRate / 1024
            Else
                If THEEXEC.enableWord("DownSample_IProfile") = True Then
                    SampleSize_Ratio = HexVsMaxSampleSize / DownSampleRatio
                    RealRate = SampleSize_Ratio / measuretime
                ElseIf SampleRate > 0 Then
                    RealRate = SampleRate
                Else
                    RealRate = HexVsMaxSampleSize / measuretime
                End If
                'RealRate = HexVsMaxSampleSize / measuretime
                If RealRate < (HexVsMaxSampleRate / 1024) Then
                    SampleRate = HexVsMaxSampleRate / 1024
                    sampleSize = SampleRate * measuretime
                Else
                
                    For i = 1 To 1024
                        If RealRate > (HexVsMaxSampleRate / i) Then
                            If i = 1 Then
                                SampleRate = (HexVsMaxSampleRate / i)
                                sampleSize = SampleRate * measuretime
                                Exit For
                            Else
                                prediff = Abs((HexVsMaxSampleRate / (i - 1)) - RealRate)
                                posdiff = Abs((HexVsMaxSampleRate / i) - RealRate)
                                If prediff > posdiff Then
                                    SampleRate = (HexVsMaxSampleRate / i)
                                    sampleSize = SampleRate * measuretime
                                Else
                                    SampleRate = (HexVsMaxSampleRate / (i - 1))
                                    sampleSize = SampleRate * measuretime
                                End If ''Sample size and sample rate should meet the spec.
                                If (SampleRate <= HexVsMaxSampleRate) And (sampleSize <= HexVsMaxSampleSize) Then Exit For
                            End If
                        Else
                        End If
                    Next i
                End If
                
            End If
        Case glbConstIns_VHDVS
            If theexec.enableWord("DownSample_IProfile") = True Then If DownSampleRatio > 2 Then DownSampleRatio = 2
            
            If measuretime > UVSMaxtime Then
                sampleSize = UVSMaxSampleSize
                SampleRate = UVSMaxSampleRate / (2 ^ 12)
            Else
                    If THEEXEC.enableWord("DownSample_IProfile") = True Then
                        SampleSize_Ratio = UVSMaxSampleSize / DownSampleRatio
                        RealRate = SampleSize_Ratio / measuretime
                    ElseIf SampleRate > 0 Then
                        RealRate = SampleRate
                    Else
                        RealRate = UVSMaxSampleSize / measuretime
                    End If
                
                If RealRate < (UVSMaxSampleRate / (2 ^ 12)) Then
                    SampleRate = UVSMaxSampleRate / (2 ^ 12)
                    sampleSize = SampleRate * measuretime
                Else
                    For i = 0 To 12
                        If RealRate > (UVSMaxSampleRate / (2 ^ i)) Then
                            If i = 1 Then
                                SampleRate = (UVSMaxSampleRate / (2 ^ i))
                                sampleSize = SampleRate * measuretime
                                Exit For
                            Else
                                prediff = Abs((UVSMaxSampleRate / 2 ^ (i - 1)) - RealRate)
                                posdiff = Abs((UVSMaxSampleRate / (2 ^ i)) - RealRate)
                                If prediff > posdiff Then
                                    SampleRate = (UVSMaxSampleRate / (2 ^ i))
                                    sampleSize = SampleRate * measuretime
                                Else
                                    SampleRate = (UVSMaxSampleRate / 2 ^ (i - 1))
                                    sampleSize = SampleRate * measuretime
                                End If ''Sample size and sample rate should meet the spec.
                                If (SampleRate <= UVSMaxSampleRate) And (sampleSize <= UVSMaxSampleSize) Then Exit For
                            End If
                        End If
                    Next i
                End If
            End If
        Case glbConstIns_VSM
            If measuretime > VSMMaxtime Then
                sampleSize = VSMMaxSampleSize
                SampleRate = VSMMaxSampleRate / (2 ^ 12)
            Else
                If THEEXEC.enableWord("DownSample_IProfile") = True Then
                    SampleSize_Ratio = VSMMaxSampleSize / DownSampleRatio
                    RealRate = SampleSize_Ratio / measuretime
                ElseIf SampleRate > 0 Then
                    RealRate = SampleRate
                Else
                    RealRate = VSMMaxSampleSize / measuretime
                End If
                If RealRate < (VSMMaxSampleRate / (2 ^ 12)) Then
                    SampleRate = VSMMaxSampleRate / (2 ^ 12)
                    sampleSize = SampleRate * measuretime
                Else
                    For i = 0 To 12
                        If RealRate > (VSMMaxSampleRate / (2 ^ i)) Then
                            If i = 1 Then
                                SampleRate = (VSMMaxSampleRate / (2 ^ i))
                                sampleSize = SampleRate * measuretime
                                Exit For
                            Else
                                prediff = Abs((VSMMaxSampleRate / 2 ^ (i - 1)) - RealRate)
                                posdiff = Abs((VSMMaxSampleRate / (2 ^ i)) - RealRate)
                                If prediff > posdiff Then
                                    SampleRate = (VSMMaxSampleRate / (2 ^ i))
                                    sampleSize = SampleRate * measuretime
                                Else
                                    SampleRate = (VSMMaxSampleRate / 2 ^ (i - 1))
                                    sampleSize = SampleRate * measuretime
                                End If ''Sample size and sample rate should meet the spec.
                                If (SampleRate <= VSMMaxSampleRate) And (sampleSize <= VSMMaxSampleSize) Then Exit For
                            End If
                        End If
                    Next i
                End If
            End If
            
            
        Case glbConstIns_VS5A, glbConstIns_VS800MA
            
            If measuretime > VSMaxtime Then
                sampleSize = VSMaxSampleSize
                SampleRate = VSMaxSampleRate / (2 ^ 12)
            Else
                    If THEEXEC.enableWord("DownSample_IProfile") = True Then
                        SampleSize_Ratio = VSMaxSampleSize / DownSampleRatio
                        RealRate = SampleSize_Ratio / measuretime
                    ElseIf SampleRate > 0 Then
                        RealRate = SampleRate
                    Else
                        RealRate = VSMaxSampleSize / measuretime
                    
                    End If
                
                If RealRate < (VSMaxSampleRate / (2 ^ 12)) Then
                    SampleRate = VSMaxSampleRate / (2 ^ 12)
                    sampleSize = SampleRate * measuretime
                Else
                    For i = 0 To 12
                        ''' VSM sample rate can't equal 25000 from HELP
                        If (VSMaxSampleRate / 2 ^ i) <> 25000 Then
                            If RealRate > (VSMaxSampleRate / (2 ^ i)) Then
                                If i = 1 Then
                                    SampleRate = (VSMaxSampleRate / (2 ^ i))
                                    sampleSize = SampleRate * measuretime
                                    Exit For
                                Else
                                    If (VSMaxSampleRate / 2 ^ (i - 1)) = 25000 And i >= 2 Then
                                        prediff = Abs((VSMaxSampleRate / 2 ^ (i - 2)) - RealRate)
                                    Else
                                        prediff = Abs((VSMaxSampleRate / 2 ^ (i - 1)) - RealRate)
                                    End If
                                    posdiff = Abs((VSMaxSampleRate / (2 ^ i)) - RealRate)
            
                                    If prediff > posdiff Then
                                        SampleRate = (VSMaxSampleRate / (2 ^ i))
                                        sampleSize = SampleRate * measuretime
                                    Else
                                        SampleRate = (VSMaxSampleRate / 2 ^ (i - 1))
                                        sampleSize = SampleRate * measuretime
                                    End If
                                    If (SampleRate <= VSMaxSampleRate) And (sampleSize <= VSMaxSampleSize) Then Exit For
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
            
    End Select
    
    sampleSize = Ceiling(sampleSize)
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "ProfileAutoResolution")
    If AbortTest Then Exit Function Else Resume Next
End Function


Function Bin2Dec_rev_Fractional(sMyBin As String) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec_rev_Fractional = Bin2Dec_rev_Fractional + mid(sMyBin, iLen - x + 1, 1) * 2 ^ (-x - 1)
    Next
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Bin2Dec_rev_Fractional") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Dec2Bin_str(ByVal n As Long, ByRef BinArray_string As String, bitCount As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim i As Integer, j As Integer
    Dim Element_Amount As Integer
    Dim Count As Integer
    Dim BinArray() As Long
    BinArray_string = vbNullString
    ReDim BinArray(bitCount) As Long
    
    '               01101
    ' BinArray(4) 1
    ' BinArray(3) 0
    ' BinArray(2) 1
    ' BinArray(1) 1
    ' BinArray(0) 0

    Element_Amount = UBound(BinArray)
    If n > (2 ^ (Element_Amount + 1) - 1) Then
        n = 0
        theexec.Datalog.WriteComment "Error(Dec2Bin): Overange for " & n
    End If

    For j = 0 To Element_Amount
        BinArray(j) = 0
    Next j

    'If n < 0 Then MsgBox ("Warning(Dec2Bin)!!! Decimal Number should be positive integer")
    If n < 0 Then
        theexec.Datalog.WriteComment " The input vlaue of (Dec2Bin) is negative, so we enforce it as 0 to prevent from error alarm."
        n = 0
    End If
    
    i = 0
    Do Until n = 0
        If (i > Element_Amount) Then theexec.Datalog.WriteComment "Warning (Dec2Bin)!!! Decimal " & n & " is over-range (>" & i & "bit)"
        If (n Mod 2) Then
            BinArray(Element_Amount - i) = 1
        Else
            BinArray(Element_Amount - i) = 0
        End If
        n = Int(n / 2)
        i = i + 1
        
    Loop
    For i = 0 To UBound(BinArray)
        BinArray_string = BinArray_string & BinArray(UBound(BinArray) - i)
    Next i

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Dec2Bin_str") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function ShmooEndFunction() As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If theexec.DevChar.Setups.IsRunning = True Then
        Dim site As Variant
        Dim SetupName As String
        Dim X_RangeFrom As Double
        Dim Y_RangeFrom As Double
        
        SetupName = theexec.DevChar.Setups.ActiveSetupName
        If Not ((theexec.DevChar.Results(SetupName).StartTime Like "1/1/0001*" Or theexec.DevChar.Results(SetupName).StartTime Like "0001/1/1*")) Then
            gl_Flag_HardIP_Characterization_1stRun = False
            With theexec.DevChar.Setups(SetupName)
                If .Shmoo.Axes.Count > 1 Then
                    X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
                    Y_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.from
                    For Each site In theexec.sites ''20181101 current point need site value
                        XVal = theexec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).value
                        YVal = theexec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value
                    Next site
                    If XVal = X_RangeFrom And YVal = Y_RangeFrom Then
                        gl_flag_end_shmoo = False
                    End If
                    If gl_flag_end_shmoo = True Then
                        ShmooEndFunction = True
                    End If
                Else
                    X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
                    For Each site In theexec.sites ''20181101 current point need site value
                        XVal = theexec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).value
                    Next site
'                    If XVal = X_RangeFrom Then
'                        gl_flag_end_shmoo = False
'                    End If
                    If gl_flag_end_shmoo = True Then
                        ShmooEndFunction = True
                    End If
                End If
            End With
        Else
           gl_Flag_HardIP_Characterization_1stRun = True
           gl_flag_end_shmoo = False
        End If
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "ShmooEndFunction") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''20230223: Modidfied to put HardIP variable "gl_GetInstrumentType_Dic" and "gl_GetInstrument_Dic" in.
''20230223: Modidfied add gl_dicPowerPinIndex to faster find PowerPin_range_ary index.
' [20230801][All][Oliver] Modidfied gl_GetInstrumentType_Dic not save :C pin merge quantity.
' [20230809][All][Oliver] add save global all power and core power DCVS/DCVI pin strings.
' [20240110][All][Tank] add store min Ifold
Public Function Get_Channel_Type()
    
    Dim funcName As String:: funcName = "Get_Channel_Type"
    
    Dim i As Long
    Dim j As Long
    Dim indx As Long
    Dim Count As Long: Count = 0
    Dim max_row As Long
    Dim max_col As Long
    Dim Pin_row As Long
    Dim Pin_col As Long
'''    Dim Type_row As Long
    Dim Type_col As Long
    
    Dim pin_name As String: pin_name = "Pin Name"
    Dim Type_Name As String: Type_Name = "Type"
    Dim PinName_Temp As String
    Dim PinMerged_Temp As String
        
    Dim sheetNames() As String
    
    Dim arr_content() As Variant
    Dim chanmap_content() As Variant
    
    Dim HexVS_RangeList_DBL() As Double
    
    Dim CRWT As New Class_CRWT ' Add for Current range wait time
    
    'Dim Dict_MergedDiffModule As New Scripting.Dictionary
    Dim Dict_MergedDiffModule As New Dictionary
    ReDim HexVS_RangeList_DBL(3)
    HexVS_RangeList_DBL(0) = 0.01: HexVS_RangeList_DBL(1) = 0.1: HexVS_RangeList_DBL(2) = 1: HexVS_RangeList_DBL(3) = 15
       
    Dim AllPower() As String
    Dim AllPowercnt As Long
    Dim corePower() As String
    Dim CorePowercnt As Long
    Dim k As Long
    
    On Error GoTo errHandler
       
    If Flag_GetChannelType = False Then
        sheetNames = theexec.job.GetSheetNamesOfType(DMGR_SHEET_TYPE_CHANMAP)
        Application.ScreenUpdating = False
    
        For indx = 0 To UBound(sheetNames)
            '''Choose the current Channel Map
            If LCase(sheetNames(indx)) = CurrentChannelMap Then
               Worksheets(sheetNames(indx)).Activate
               max_row = Worksheets(sheetNames(indx)).UsedRange.Rows.Count
               max_col = Worksheets(sheetNames(indx)).UsedRange.Columns.Count
               chanmap_content = Worksheets(sheetNames(indx)).range(Cells(1, 1), Cells(max_row, max_col)).value
    
               For i = 1 To max_row
                   For j = 1 To max_col - 1
                       If chanmap_content(i, j) <> "" Then
                           If chanmap_content(i, j) = pin_name Then '''Get the position of "Pin Name" from channel map
                               Pin_row = i
                               Pin_col = j
                           ElseIf chanmap_content(i, j) = Type_Name Then '''Get the position of "Type" from channel map
    '''                            Type_row = i
                               Type_col = j
                               Exit For
                           End If
                       End If
                   Next j
               Next i
               
               Exit For
            End If
            
        Next indx
        
        ReDim Pin_range_ary(max_row - Pin_row - 1)
        For i = Pin_row + 1 To max_row
            For j = Pin_col To max_col - 1
                If j = Pin_col Then '''Get the "Pin Name" from array in channel map
                    Pin_range_ary(i - (Pin_row + 1)).PinName = chanmap_content(i, j)
                    If InStr(UCase(chanmap_content(i, j)), ":C") > 0 Then
                        chanmap_content(i, j) = Replace(UCase(chanmap_content(i, j)), ":C", "")
                        If Not Dict_MergedDiffModule.Exists(chanmap_content(i, j)) Then Dict_MergedDiffModule.Add UCase(chanmap_content(i, j)), 0
                    ElseIf InStr(UCase(chanmap_content(i, j)), ":M") > 0 Then
                        Call Print_Error_Message(Warning_Info, LIB_Common, "Get_Channel_Type", "Channel Map pin includes :M pin!!!")
                    End If
                ElseIf j = Type_col Then '''Get the MergeType from array in channel map
                    Pin_range_ary(i - (Pin_row + 1)).MergeType = chanmap_content(i, j)
                End If
            Next j
        Next i
        j = 0
        Application.ScreenUpdating = True
        
        For i = 0 To UBound(Pin_range_ary)
            ''Check if it is merged with other modules
''            If Pin_range_ary(i).PinName <> "" And Pin_range_ary(i).MergeType <> "N/C" Then
            If Pin_range_ary(i).PinName <> "" Then
                PinName_Temp = UCase(Pin_range_ary(i).PinName)
                PinMerged_Temp = Pin_range_ary(i).MergeType
                If Dict_MergedDiffModule.Exists(PinName_Temp) Then
                    Dict_MergedDiffModule.Remove PinName_Temp
                    Dict_MergedDiffModule.Add PinName_Temp, PinMerged_Temp
                End If
                If InStr(PinName_Temp, ":C") > 0 Then
                    PinName_Temp = Replace(PinName_Temp, ":C", "")
                    If Dict_MergedDiffModule.Exists(PinName_Temp) Then
                        PinMerged_Temp = "DCVSMerged" & CStr(CLng(Find_Channel_MergeCase(Dict_MergedDiffModule.item(PinName_Temp))) + CLng(Find_Channel_MergeCase(PinMerged_Temp)))
                        Dict_MergedDiffModule.Remove PinName_Temp
                        Dict_MergedDiffModule.Add PinName_Temp, PinMerged_Temp
                    End If
                End If
            End If
            Call Store_Special_pins_pair(Pin_range_ary(i).PinName) '20240702 michael add for store special pin pair
        Next i
        
        Dim sTempPinName As String
        For i = 0 To UBound(Pin_range_ary)
            '''Do if PinName is not emtpy and MergeType is not N/C pin in channel map sheet
            If Pin_range_ary(i).PinName <> "" And Pin_range_ary(i).MergeType <> "N/C" Then
''            If Pin_range_ary(i).PinName <> "" Then
                If InStr(UCase(Pin_range_ary(i).PinName), ":C") > 0 Or InStr(UCase(Pin_range_ary(i).PinName), ":M") > 0 Then
                Else
                
                    '========================================
                    sTempPinName = LCase(Pin_range_ary(i).PinName)

                    If gl_GetInstrumentType_Dic.Exists(sTempPinName) = False Then
                        If Dict_MergedDiffModule.Exists(UCase(sTempPinName)) Then
                            Call gl_GetInstrumentType_Dic.Add(sTempPinName, Dict_MergedDiffModule(Pin_range_ary(i).PinName))
                        Else
                            Call gl_GetInstrumentType_Dic.Add(sTempPinName, Pin_range_ary(i).MergeType)     'store PinName ,channelType (ex: VDDIO12_GRP0, DCVSMerged4)
                        End If
                    End If

                    If gl_GetInstrument_Dic.Exists(sTempPinName) = False Then
                        Call gl_GetInstrument_Dic.Add(sTempPinName, GetInstrument(Pin_range_ary(i).PinName, 0))     'store PinName ,Instrument (ex: VDDIO12_GRP0, vhdvs)
                    End If
                    '========================================
                    
                    Pin_range_ary(i).PinMapType = theexec.DataManager.PinType(Pin_range_ary(i).PinName)
                    '''Do if PinMapType is Power in PinMap sheet
                    If Pin_range_ary(i).PinMapType Like "*Power*" Then
                        
                        ReDim Preserve PowerPin_range_ary(j)
                        
                        PowerPin_range_ary(j).PinName = Pin_range_ary(i).PinName
                        
                        If gl_dicPowerPinIndex.Exists(LCase(PowerPin_range_ary(j).PinName)) Then    '20230112 Serch PowerPin_range_ary use dic "PinName, index"
                            gl_dicPowerPinIndex.Remove (LCase(PowerPin_range_ary(j).PinName))
                        Else
                        End If

                        gl_dicPowerPinIndex.Add LCase(PowerPin_range_ary(j).PinName), j

                        '''Check if it is merged with other modules
                        If Dict_MergedDiffModule.Exists(UCase(PowerPin_range_ary(j).PinName)) Then
                        
                            PowerPin_range_ary(j).MergeType = Dict_MergedDiffModule.item(UCase(PowerPin_range_ary(j).PinName))
                        Else
                            PowerPin_range_ary(j).MergeType = Pin_range_ary(i).MergeType
                        End If
                        
                        PowerPin_range_ary(j).PinMapType = Pin_range_ary(i).PinMapType
                        '''Get the instrument type, i.g., "DCVS", "DCVSMergedN" and "DCVI"...
                        PowerPin_range_ary(j).ChanMapType = GetInstrument(PowerPin_range_ary(j).PinName, 0)
                        '''Get merged value from the the instrument type, i.g., "DCVS" => 1, "DCVSANALOGMUX_OUTMergedN" => N
                        PowerPin_range_ary(j).MergedN = Find_Channel_MergeCase(PowerPin_range_ary(j).MergeType)
                        '''Get the initial current range value from HW
                        If Pin_range_ary(i).MergeType Like "*DCVS*" Then
                            PowerPin_range_ary(j).Init_CurrentRange = TheHdw.DCVS.Pins(PowerPin_range_ary(j).PinName).CurrentRange.value
                        
                        ElseIf Pin_range_ary(i).MergeType Like "*DCVI*" Then
                            PowerPin_range_ary(j).Init_CurrentRange = TheHdw.DCVI.Pins(PowerPin_range_ary(j).PinName).CurrentRange.value
            
                        End If
                        '''Get the current range list from HW
                        If LCase(PowerPin_range_ary(j).ChanMapType) = "vhdvs" Or LCase(PowerPin_range_ary(j).ChanMapType) = "vsm" _
                                Or LCase(PowerPin_range_ary(j).ChanMapType) = "vs-5a" Or LCase(PowerPin_range_ary(j).ChanMapType) = "vs-800ma" Then
                            PowerPin_range_ary(j).Range_List = TheHdw.DCVS.Pins(PowerPin_range_ary(j).PinName).CurrentRange.list
                        ElseIf LCase(PowerPin_range_ary(j).ChanMapType) = "hexvs" Then
                            PowerPin_range_ary(j).Range_List = HexVS_RangeList_DBL ''''''carter 0122
                            If PowerPin_range_ary(j).MergeType Like "*DCVSMerged*" Then
                                ReDim Preserve PowerPin_range_ary(j).Range_List(4)
                                PowerPin_range_ary(j).Range_List(4) = TheHdw.DCVS.Pins(PowerPin_range_ary(j).PinName).Meter.CurrentRange
                            End If
                            
                        ElseIf LCase(PowerPin_range_ary(j).ChanMapType) = "dc-07" Then
                            PowerPin_range_ary(j).Range_List = TheHdw.DCVI.Pins(PowerPin_range_ary(j).PinName).CurrentRange.list
                        
                        End If
                        '''Get the waittime list for each current range list, respectively
                        PowerPin_range_ary(j).WaitTime_List = Find_Range_WaitTime(PowerPin_range_ary(j).Range_List, PowerPin_range_ary(j).ChanMapType, PowerPin_range_ary(j).MergedN, j)
                        '''Get the waittime list for each current range list, respectively
                        PowerPin_range_ary(j).Accuracy_List = Find_Range_Accuracy(PowerPin_range_ary(j).Range_List, PowerPin_range_ary(j).ChanMapType, PowerPin_range_ary(j).MergedN)
                        
                        '==== Get DCVS pin min IFold limit ====
                        PowerPin_range_ary(j).MinIFoldLimit = GetMinIFoldLimit(PowerPin_range_ary(j).PinName, PowerPin_range_ary(j).Range_List(0))
                        '==== Get DCVS pin min IFold limit ====
                         
                        j = j + 1
                                   
                    End If
                    
                End If
            
            End If
            
        Next i
        '''Do once in the 1st touch down
        Flag_GetChannelType = True

        ALL_Power_DCVS_pins = vbNullString
        ALL_Power_DCVI_pins = vbNullString
        Core_Power_DCVS_pins = vbNullString
        Core_Power_DCVI_pins = vbNullString

        If AllPowerPinlist <> "" Then
            Call SortAllPinInstrumentType(AllPowerPinlist, ALL_Power_DCVS_pins, ALL_Power_DCVI_pins)
        Else
           Call Print_Error_Message(Warning_Info, "LIB_Common", funcName, "PinMap not have All Power group!!!")
        End If
        If CorePowerPinlist <> "" Then
            Call SortAllPinInstrumentType(AllPowerPinlist, Core_Power_DCVS_pins, Core_Power_DCVI_pins)
        Else
           Call Print_Error_Message(Warning_Info, "LIB_Common", funcName, "PinMap not have core Power group!!!")
        End If
    End If
    If Not F_p2pWT Then F_p2pWT = CRWT.LoadCRWTToPowerPinAry("TuneCRWTResult", "CRWT_Org")
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "Get_Channel_Type")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Find_Channel_MergeCase(MergeType As String) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "Find_Channel_MergeCase"
    
    
    Find_Channel_MergeCase = 1
    
    If MergeType Like "*DCVSMerged*" Then
        Find_Channel_MergeCase = Replace(MergeType, "DCVSMerged", "")
    
    ElseIf MergeType Like "*DCVIMerged*" Then
        Find_Channel_MergeCase = 2
    
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Find_Channel_MergeCase") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


'[20240312][T-All][Clyde] Updated to support new power up/down function
Public Function Get_PowerSeq()
On Error GoTo errHandler
    Dim i As Long
    Dim j As Long
    Dim tmpSeqNum As Integer
    Dim gloSpecList() As String
    
    functionName = "Get_PowerSeq"
    
    If Flag_GetPowerSeq = False Then
        ' Read Global Spec list
        gloSpecList = theexec.Specs.Globals.list
        nonPowerSeqNum.RemoveAll
        maxSeqNum = 0
        If Len(Join(gloSpecList)) = 0 Then
            Call Print_Error_Message(Error_Info, moduleName, functionName, "No Global Spec, please check!!")
        Else
            For i = 0 To UBound(gloSpecList)
                For j = 0 To UBound(PowerPin_range_ary)
                    If gloSpecList(i) Like PowerPin_range_ary(j).PinName & PowerUpSeqName & "*" Then
                        tmpSeqNum = theexec.Specs.Globals(gloSpecList(i)).ContextValue
                        If maxSeqNum < tmpSeqNum And tmpSeqNum <> 99 Then maxSeqNum = tmpSeqNum
                        PowerPin_range_ary(j).PowerSeq = tmpSeqNum
                        Exit For
                    ElseIf gloSpecList(i) Like PowerPin_range_ary(j).PinName & PowerDownSeqName & "*" Then
                        tmpSeqNum = theexec.Specs.Globals(gloSpecList(i)).ContextValue
                        If maxSeqNum < tmpSeqNum And tmpSeqNum <> 99 Then maxSeqNum = tmpSeqNum
                        PowerPin_range_ary(j).PowerDownSeq = theexec.Specs.Globals(gloSpecList(i)).ContextValue
                        Exit For
                    End If
                Next j
                If gloSpecList(i) Like "*" & PowerDownSeqName & "*" Or gloSpecList(i) Like "*" & PowerUpSeqName & "*" Then
                    If tmpSeqNum = 0 Then
                        tmpSeqNum = theexec.Specs.Globals(gloSpecList(i)).ContextValue
                        If maxSeqNum < tmpSeqNum And tmpSeqNum <> 99 Then maxSeqNum = tmpSeqNum
                        nonPowerSeqNum.Add UCase(gloSpecList(i)), tmpSeqNum
                    End If
                End If
                tmpSeqNum = 0
            Next i
            Flag_GetPowerSeq = True
        End If
    End If
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20240311][T-All][Clyde] Updated Power up for TTR
Public Function PowerUpDown_Process(WaitConnectTime As Double, DebugFlag As Boolean, isUp As Boolean)
On Error GoTo errHandler
    Dim i As Integer
    Dim j As Integer
    Dim nwirePortPlist As New PinList
    Dim pwrSeq As PowerSeqeuce
    Dim printUpDown As String
    
    If isUp Then
        pwrSeq = pwrUpSeq
        printUpDown = "power up"
    Else
        pwrSeq = pwrDownSeq
        printUpDown = "power down"
    End If
    
    functionName = "PowerUpDown_Process"
    
    For i = 0 To maxSeqNum
        '------------------------------------power pin sequence
        If isUp Then
            If pwrSeq.PowerSeqPin(i) <> "" Then
                ''power up
                theexec.Datalog.WriteComment vbCrLf & "print: " & printUpDown & " action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                PowerOnOff_I_Meter_Parallel pwrSeq.PowerSeqPin(i), WaitConnectTime, WaitConnectTime, i, isUp, DebugFlag
            End If
            '------------------------------------Support multiple nWire port 20170503
            If Len(Join(pwrSeq.nWirePort)) <> 0 Then
                For j = 0 To UBound(pwrSeq.nWirePort)
                    If (pwrSeq.nWireSeq(j) = i) And (pwrSeq.nWirePort(j) <> "") Then
                        SetUpPowerUpDownFRCState pwrSeq.nWirePort(j), i, isUp, DebugFlag
                    End If
                Next j
            End If
            '------------------------------------Support I/O init_H
            If Len(Join(pwrSeq.IOHPin)) <> 0 Then
                For j = 0 To UBound(pwrSeq.IOHPin)
                    If pwrSeq.IOHSeq(j) = i Then
                        Call SetUpPowerUpDownIOState(pwrSeq.IOHPin(j), pwrSeq.IOHSeq(j), chInitHi, isUp)
                    End If
                Next j
            End If
            '------------------------------------Support I/O init_L
            If Len(Join(pwrSeq.IOLPin)) <> 0 Then
                For j = 0 To UBound(pwrSeq.IOLPin)
                    If pwrSeq.IOLSeq(j) = i Then
                        Call SetUpPowerUpDownIOState(pwrSeq.IOLPin(j), pwrSeq.IOLSeq(j), chInitLo, isUp)
                    End If
                Next j
            End If
            '------------------------------------Support I/O init_HIZ
            If Len(Join(pwrSeq.IOHZPin)) <> 0 Then
                For j = 0 To UBound(pwrSeq.IOHZPin)
                    If pwrSeq.IOHZSeq(j) = i Then
                        Call SetUpPowerUpDownIOState(pwrSeq.IOHZPin(j), pwrSeq.IOHZSeq(j), chInitoff, isUp)
                    End If
                Next j
            End If
        Else
            '------------------------------------Support I/O init_HIZ
            If Len(Join(pwrSeq.IOHZPin)) <> 0 Then
                For j = 0 To UBound(pwrSeq.IOHZPin)
                    If pwrSeq.IOHZSeq(j) = i Then
                        Call SetUpPowerUpDownIOState(pwrSeq.IOHZPin(j), pwrSeq.IOHZSeq(j), chInitoff, isUp)
                    End If
                Next j
            End If
            '------------------------------------Support I/O init_L
            If Len(Join(pwrSeq.IOLPin)) <> 0 Then
                For j = 0 To UBound(pwrSeq.IOLPin)
                    If pwrSeq.IOLSeq(j) = i Then
                        Call SetUpPowerUpDownIOState(pwrSeq.IOLPin(j), pwrSeq.IOLSeq(j), chInitLo, isUp)
                    End If
                Next j
            End If
            '------------------------------------Support I/O init_H
            If Len(Join(pwrSeq.IOHPin)) <> 0 Then
                For j = 0 To UBound(pwrSeq.IOHPin)
                    If pwrSeq.IOHSeq(j) = i Then
                        Call SetUpPowerUpDownIOState(pwrSeq.IOHPin(j), pwrSeq.IOHSeq(j), chInitHi, isUp)
                    End If
                Next j
            End If
            '------------------------------------Support multiple nWire port 20170503
            If Len(Join(pwrSeq.nWirePort)) <> 0 Then
                For j = 0 To UBound(pwrSeq.nWirePort)
                    If (pwrSeq.nWireSeq(j) = i) And (pwrSeq.nWirePort(j) <> "") Then
                        SetUpPowerUpDownFRCState pwrSeq.nWirePort(j), i, isUp, DebugFlag
                    End If
                Next j
            End If
            If pwrSeq.PowerSeqPin(i) <> "" Then
                ''power down
                theexec.Datalog.WriteComment vbCrLf & "print: " & printUpDown & " action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                PowerOnOff_I_Meter_Parallel pwrSeq.PowerSeqPin(i), WaitConnectTime, WaitConnectTime, i, isUp, DebugFlag
            End If
        End If
    Next i
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20231108][All][Oliver] modify to not generate CZMappingTable_List
Public Function CheckPatAndCZmappingTable()
    
    On Error GoTo errHandler

    Dim sFuncName As String:: sFuncName = "CheckPatAndCZmappingTable"
    
    Dim i As Long
    Dim j As Long
    
    Dim maxcol As Long
    Dim MaxRow As Long
    Dim cz_pat_counter As Long
    Dim patset_all_counter As Long
    Dim patset_all_index_x As Long
    Dim patset_all_index_y As Long
    

    Dim temp_str As String
    Dim filename As String
    Dim cz_sheetname As String
    Dim cz_pat_header As String
    Dim patset_all_header As String
    
    Dim cz_pat_module() As String
    Dim patset_all_module() As String
    Dim pattern_sheetname() As String
    Dim patset_module_nofound() As String
    
    Dim patset_all_Ary() As Variant
    Dim CZ_mappingTable_Ary() As Variant
    
    Dim Flag_pat_dssc As Boolean
    Dim Flag_pat_found As Boolean
    Dim Flag_pat_found_final As Boolean
    Dim Flag_patset_all As Boolean
    Dim Flag_FindPatternName As Boolean
    Dim PatternName_Row As Long
    
    If gl_isPatAndCZmappingTable = False Then
        If WorksheetExists("PatSets_All", False) = False Then
            MsgBox "Sheet PatSets_All is not exist!!!", vbOKOnly + vbCritical, "Error"
            Exit Function
        End If
        
        patset_all_header = "Pattern Set"
        pattern_sheetname = theexec.job.GetSheetNamesOfType(DMGR_SHEET_TYPE_PATTERNSETSHEET)
        ''********Store PatSets_All contents********
        For i = 0 To UBound(pattern_sheetname)
            If pattern_sheetname(i) = "PatSets_All" Then
                Worksheets(pattern_sheetname(i)).Activate
                MaxRow = Worksheets(pattern_sheetname(i)).UsedRange.Rows.Count
                maxcol = Worksheets(pattern_sheetname(i)).UsedRange.Columns.Count
                patset_all_Ary() = Worksheets(pattern_sheetname(i)).range(Cells(1, 1), Cells(MaxRow, maxcol)).value
                Exit For
            End If
        Next i
        ''********Store PatSets_All contents********
        
        ''--------Save the pat index which related to the specified pattern header "Pattern Set"--------
        For i = 1 To MaxRow
            For j = 1 To maxcol
                If patset_all_Ary(i, j) = patset_all_header Then
                    patset_all_index_x = i
                    patset_all_index_y = j
                    Flag_patset_all = True
                    Exit For
                End If
            Next j
            If Flag_patset_all Then Exit For
        Next i
        ''--------Save the pat index which related to the specified pattern header "Pattern Set"--------
        
        ''--------Store the pattenr module for specified pat index--------
        patset_all_counter = 0
        For i = patset_all_index_x + 1 To MaxRow
            If UCase(patset_all_Ary(i, patset_all_index_y)) Like "*SRMDSSC*" Then
                Flag_pat_dssc = True
                ReDim Preserve patset_all_module(patset_all_counter)
                patset_all_module(patset_all_counter) = patset_all_Ary(i, patset_all_index_y)
                patset_all_counter = patset_all_counter + 1
            End If
        Next i
        ''--------Store the pattenr module for specified pat index--------
        
''**    ******If there is no dssc pattern in the "PatSets_All", we will not check the CZMappingTable********
        If Flag_pat_dssc Then
        
            cz_sheetname = "DSSCMappingTable_" & UCase(currentJobName)
''            cz_sheetname = "CZ_DSSCmappingtable_CP1"
            If WorksheetExists(cz_sheetname, False) = False Then
                MsgBox "Sheet " & cz_sheetname & " is not exist!!!", vbOKOnly + vbCritical, "Error"
                Exit Function
            End If
            
            Worksheets(cz_sheetname).Activate
            
            MaxRow = Worksheets(cz_sheetname).UsedRange.Rows.Count
            maxcol = Worksheets(cz_sheetname).UsedRange.Columns.Count
            ReDim CZ_mappingTable_Ary(MaxRow, maxcol)
            
            cz_pat_header = "Pattern Name"
            ''Store CZ_mappingTable contents
            CZ_mappingTable_Ary() = Worksheets(cz_sheetname).range(Cells(1, 1), Cells(MaxRow, maxcol)).value
            
            ''********Save the pat index which related to the specified pattern header "Pattern Name"********
            cz_pat_counter = 0
            For i = 1 To MaxRow
                If Flag_FindPatternName Then
                    If CZ_mappingTable_Ary(i, 1) <> "" Then
                        ReDim Preserve cz_pat_module(cz_pat_counter)
                        ''--------Store the pattenr module--------
                        cz_pat_module(cz_pat_counter) = CZ_mappingTable_Ary(i, 1)
                        ''--------Store the pattern module for specified pat index--------
                        cz_pat_counter = cz_pat_counter + 1
                    Else
                    '''If TempAry_SheetContent is empty, setup flag "Flag_FindPatternName" as False
                        Flag_FindPatternName = False
                    End If
                Else
                    '''Find the Header "Pattern Name", record the row value and setup flag "Flag_FindPatternName" as True
                    If CZ_mappingTable_Ary(i, 1) Like cz_pat_header Then
                        PatternName_Row = i
                        Flag_FindPatternName = True
                    End If
                End If
            Next i
            ''--------Store the pattern module for specified pat index--------
            
            
            ''********Compare the pattern module between CZ_MappingTable_* and PatSets_All********
            patset_all_counter = 0
            Flag_pat_found = True
            Flag_pat_found_final = True
            For i = 0 To UBound(patset_all_module)
                Flag_pat_found = False
''                If UCase(patset_all_module(i)) Like "*DSSC*" Then
                For j = 0 To UBound(cz_pat_module)
                    If LCase(patset_all_module(i)) = LCase(cz_pat_module(j)) Then
                        Flag_pat_found = True
                    Else
                        Flag_pat_found = False
                    End If
                    If Flag_pat_found Then Exit For
                Next j
                
                If Flag_pat_found = False Then
                    ReDim Preserve patset_module_nofound(patset_all_counter)
                    patset_module_nofound(patset_all_counter) = patset_all_module(i)
                    patset_all_counter = patset_all_counter + 1
                    Flag_pat_found_final = Flag_pat_found_final And Flag_pat_found
                End If
                    
''                End If
            Next i
            ''********Compare the pattern module between CZ_MappingTable_* and PatSets_All********
            
            ''--------MessageBox dispalays the pattern module which is not defined in the CZMappingTable--------
            If Flag_pat_found_final = False Then
'                filename = CurDir & "\" & "CZMappingTable_List.txt"
'                Open filename For Output As #1
'                For i = 0 To UBound(patset_module_nofound)
'                    Print #1, patset_module_nofound(i) & vbCrLf
'                Next i
'                Close #1
                theexec.ErrorLogMessage "Error: Pattern Module misses in the CZmappingTable, Please check the list in the file!!!"
'                MsgBox "Please check the list in the file " & FileName, vbOKOnly, "Error: Pattern Module misses in the CZmappingTable!!!"
                Flag_CZMappingTable_Check = False
            Else
'                filename = CurDir & "\" & "CZMappingTable_List.txt"
'                Open filename For Output As #1
'                    Print #1, vbCrLf
'                Close #1
                Flag_CZMappingTable_Check = True
            End If
        
        End If
        gl_isPatAndCZmappingTable = True
    End If
'''    If Flag_pat_found_final = False Then
'''        For i = 0 To UBound(patset_module_nofound)
'''            If i = 0 Then
'''                temp_str = patset_module_nofound(i)
'''            Else
'''                temp_str = temp_str & "," & vbCrLf & patset_module_nofound(i)
'''            End If
'''
'''        Next i
'''
'''        MsgBox temp_str & vbCrLf & "Did not find in the " & cz_sheetname, vbOKOnly + vbCritical, "Error: Pattern Module misses in the CZmappingTable!!!"
'''
'''    End If
'''
'''    Flag_CZMappingTable_Check = True
    ''--------MessageBox dispalays the pattern module which is not defined in the CZMappingTable--------
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, sFuncName)
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function GetPatsFromPatSets(TestPat As String, _
                              rtnPatNames() As String, _
                              rtnPatCnt As Long, Optional patSel As Boolean = True)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    
    Dim funcName As String:: funcName = "GetPatsFromPatSets"
    Dim patset() As String
    Dim PatCnt As Long
    Dim newPatset() As String
    Dim tmpPatset() As String
    Dim tmpPatcnt As Long
    Dim Pat As Variant
    Dim tmpPat As Variant
    Dim SearchDone As Boolean

    patset = theexec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
    Do While (True)
        SearchDone = True
        Erase newPatset
        Erase tmpPatset
        For Each Pat In patset
            ''' get upper order of patset
            If Not (patSel) Then
                If LCase(Pat) Like "*.pat*" Then
                    '''the patset already contain path in one order, it is not allowed
                    If PatCnt <> 1 Then
                        theexec.ErrorLogMessage "It directly used path in patset of " & TestPat
                    Else
                        Erase rtnPatNames
                        Call appendToArray(rtnPatNames, TestPat)
                        Exit Function
                    End If
                Else
                    tmpPatset = theexec.DataManager.Raw.GetPatternsInSet(Pat, tmpPatcnt)
                    For Each tmpPat In tmpPatset
                        If LCase(tmpPat) Like "*.pat*" Then
                            ''' patset contain path and non-path, it is not allowed
                            If tmpPatcnt <> 1 Then
                                theexec.ErrorLogMessage "It directly used path in patset of " & Pat
                            Else
                                Call appendToArray(newPatset, Pat)
                            End If
                        Else
                            SearchDone = False
                            Call appendToArray(newPatset, tmpPat)
                        End If
                    Next tmpPat
                End If
            Else
                ''' get last order of patset
                If Not LCase(Pat) Like "*.pat*" Then
                    SearchDone = False
                    tmpPatset = theexec.DataManager.Raw.GetPatternsInSet(Pat, tmpPatcnt)
                    For Each tmpPat In tmpPatset
                        Call appendToArray(newPatset, tmpPat)
                    Next tmpPat
                Else
                    Call appendToArray(newPatset, Pat)
                End If
            End If
        Next Pat
        
        If SearchDone Then
            Exit Do
        End If
        Erase patset
        patset = newPatset
    Loop
    
    Erase rtnPatNames
    rtnPatNames = patset
    rtnPatCnt = UBound(rtnPatNames) + 1

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "GetPatsFromPatSets") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



Public Function appendToArray(ary() As String, val As Variant)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim tmp As Integer
    tmp = UBound(ary)
    
    If ary(UBound(ary)) <> "" Then
        ReDim Preserve ary(UBound(ary) + 1)
    End If

    ary(UBound(ary)) = val
'    ReDim Preserve ary(UBound(ary) - 1)
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    ReDim ary(0)
    Resume Next
End Function


''''20180522 update for the case with the same value on allSites
''''20200623 Substitute "eFuse_DSSC_SetupDigSrcArr_allSites" for RF
Public Function DSSC_SetupDigSrcArr_allSites(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, WaveDefArray() As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funcName As String:: funcName = "DSSC_SetupDigSrcArr_allSites"
    
    Dim inwave As New DSPWave
    Dim waveDblArray() As Double
    Dim site As Variant
    Dim WaveDef As String

    For Each site In theexec.sites.Active
        inwave.data = WaveDefArray
        inwave = inwave.ConvertDataTypeTo(DspDouble)
        waveDblArray = inwave.data
        Exit For
    Next site
    
    WaveDef = "WaveDef_" + SignalName + "_allSites"
    TheHdw.Patterns(patt).Load
    
    ''''<NOTICE> Here WaveDefArray() must be Double for this case
    theexec.WaveDefinitions.CreateWaveDefinition WaveDef, waveDblArray, True
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
    
    ''''<NOTICE> check if there is already one outside
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
    
    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef
        .sampleSize = SegmentSize
        .Amplitude = 1
        '.LoadSamples
        If LCase(glb_TesterType) = "jaguar" Then
            .LoadSettings
        End If
    End With

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "DSSC_SetupDigSrcArr_allSites") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



' [20231108][All][Carter] Restore delete function
Public Function DTS_AddStoredData() '''program start
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funcName As String:: funcName = "DTS_AddStoredData"
    '20200408 Oliver
    Dim lot_id As String, tmplot_id() As String, arr As String, FilePath As String, FileExists As String
    Dim lineFromFile As String, tmp() As String, tmp1() As String, tmplot As String, tmpocr() As String, ocrid As String
    Dim ocr_lotid, ocr_x, ocr_y, tmp_x, tmp_y As String
    
    lot_id = TheExec.Datalog.Setup.LotSetup.lotId       'get id data and convert format

    If lot_id <> "" Then ''200908 If the Lot ID is empty,it will skip parsing file
        tmplot_id = Split(lot_id, ".")
        lot_id = right(tmplot_id(0), 6)
        arr = ("ACDEFGHJKLMNPRTUVWXY")                      'convert rule
        FilePath = "C:\TOPD_ME\DTSMAP\" & lot_id & ".txt"   'open file path
        FileExists = Dir(FilePath) <> ""
        If FileExists = True Then
            Flag_DTS_function = True
            Set DictOCR = CreateObject("scripting.dictionary")
            DictOCR.RemoveAll
            Open FilePath For Input As #1
            Do While Not EOF(1)
                Line Input #1, lineFromFile
                If lineFromFile <> "" Then
                    tmp = Split(lineFromFile, ";")  'lotid : tmp(0) / ocrid : tmp(1)
                    tmpocr = Split(tmp(1), "-")     'id : tmpocr(0) / wafer : tmpocr(1) / x : tmpocr(2) / y : tmpocr(3)
                    
                    If Len(tmpocr(1)) < 2 Then      'wafer id
                        ocr_lotid = "0" + tmpocr(1)
                    Else
                        ocr_lotid = tmpocr(1)
                    End If
                    
                    If val(tmpocr(2)) - 9 > 0 And val(tmpocr(3)) - 9 > 0 Then  'x,y id
                        ocr_x = mid(arr, (val(tmpocr(2)) - 9), 1)
                        ocr_y = mid(arr, (val(tmpocr(3)) - 9), 1)
                    ElseIf val(tmpocr(2)) - 9 > 0 And val(tmpocr(3)) - 9 <= 0 Then
                        ocr_x = mid(arr, (val(tmpocr(2)) - 9), 1)
                        ocr_y = tmpocr(3)
                    ElseIf val(tmpocr(2)) - 9 <= 0 And val(tmpocr(3)) - 9 > 0 Then
                        ocr_x = tmpocr(2)
                        ocr_y = mid(arr, (val(tmpocr(3)) - 9), 1)
                    Else
                        ocr_x = tmpocr(2)
                        ocr_y = tmpocr(3)
                    End If
                    ocrid = tmpocr(0) & ocr_lotid & ocr_x & ocr_y   'last ocrid
                    
                    tmp1 = Split(tmp(0), "-")
                    'tmp1 = Replace(tmp1, " ", "")
                    If Len(tmp1(1)) < 2 Then
                        tmp1(1) = "0" & tmp1(1)
                    End If
                    If Len(tmp1(2)) < 2 Then
                        tmp_x = "0" + tmp1(2)
                    Else
                        tmp_x = tmp1(2)
                    End If
                    If Len(tmp1(3)) < 2 Then
                        tmp_y = "0" + tmp1(3)
                    Else
                        tmp_y = tmp1(3)
                    End If
                    tmplot = tmp1(0) & tmp1(1) & tmp_x & tmp_y      'last lotid
                    
                    If DictOCR.Exists(tmplot) = False Then
                        DictOCR.Add tmplot, ocrid
                    End If
                End If
            Loop
            Close #1
            Kill (FilePath)
        End If
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "DTS_AddStoredData") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231124][All][Tank] Check sheet exist or not
Public Function Parse_DSSCPat_DigSrcInfo()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim ws As Worksheet
    Dim MaxColumn As Long
    Dim MaxRow As Long
    Dim CurColumn As Long
    Dim CurRow As Long
    Dim tempStr As String
    Dim TempCount As Double
    Dim ExistPatCount As Integer
    Dim ExistTestCount As Integer
    Dim CurTestNum As Integer
    Dim sheetName As String
    
    Dim DicKey_TestCase As String
    Dim DicKey_SrcStock As String
    

    Dim TestCase_Row As Long
    Dim PatternName_Row As Long
    Dim TestCase_DigSrc_Num As Long
    
    Dim Flag_FindTestName As Boolean
    Dim Flag_FindPatternName As Boolean
    Dim vDigSrcInfo() As Variant
    
    Dim funcName As String:: funcName = "Parse_DSSCPat_DigSrcInfo"
    

    If DSSCMappingTableIsRead = False Then
        Dic_SrcStockIndex.compareMode = 1
        Dic_SrcStockIndex.RemoveAll

        Application.ScreenUpdating = False
        sheetName = "DSSCMappingTable_" & UCase(currentJobName)
        
        If GetSheetInfo(sheetName, MaxRow, MaxColumn, vDigSrcInfo) Then
            ' define array depth of patterns
            CurTestNum = 0
            ExistPatCount = 0
            TestCase_DigSrc_Num = 0
            ReDim SrcStock(ExistPatCount) As DynamicSrc
            For CurRow = 1 To MaxRow
                '''--------If "Flag_FindPatternName" as True and TempAry_SheetContent not empty--------
                If Flag_FindPatternName Then
                        If vDigSrcInfo(CurRow, 1) <> "" Then
                        '''--------Record the Pattern Name--------
                        ExistPatCount = ExistPatCount + 1
'                        If ExistPatCount = 15 Then Stop
                        ReDim Preserve SrcStock(ExistPatCount - 1) As DynamicSrc
                            SrcStock(ExistPatCount - 1).PatternName = vDigSrcInfo(CurRow, 1)
                        '''--------Record the Pattern Name--------
                        For CurColumn = 2 To MaxColumn
                            
                            If Flag_FindTestName Then
                                '''--------If "Flag_FindPatternName" as True and TempAry_SheetContent not empty--------
                                    If vDigSrcInfo(PatternName_Row, CurColumn) <> "" Then
                                    '''--------Record the TestCase, i.g., Test_Mbist_Cpu_Init_X_NV--------
                                    CurTestNum = CurTestNum + 1
                                    ReDim Preserve SrcStock(ExistPatCount - 1).TestCase(CurTestNum) As testCondition
                                        SrcStock(ExistPatCount - 1).TestCase(CurTestNum).ConditionName = vDigSrcInfo(PatternName_Row, CurColumn)
                                    '''--------Record the TestCase, i.g., Test_Mbist_Cpu_Init_X_NV--------
                                    TempCount = 0
                                    For TestCase_Row = PatternName_Row + 1 To MaxRow
                                        tempStr = vbNullString
                                            If vDigSrcInfo(TestCase_Row, CurColumn) <> "" Then
                                            '''--------Record the TestCase, i.g., S or 0--------
                                            ReDim Preserve SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStrAry(TempCount) As String
                                                SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStrAry(TempCount) = vDigSrcInfo(TestCase_Row, CurColumn)
                                            '''--------Record the TestCase, i.g., S or 0--------
                                            TempCount = TempCount + 1
                                        Else
                                            '''--------Record the TestCase, i.g., SSSSSSSSS or 0000000000--------
                                            SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStr = Join(SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStrAry, "")
                                            ''--------Record the TestCase, i.g., SSSSSSSSS or 0000000000--------
                                            SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitCount = TempCount
                                            '''********add dictionary for index of all test cases ' swlinza 20190602********
                                            DicKey_TestCase = vbNullString
                                            DicKey_TestCase = SrcStock(ExistPatCount - 1).TestCase(CurTestNum).ConditionName
                                            SrcStock(ExistPatCount - 1).TestCase_index.compareMode = 1
                                            If Not SrcStock(ExistPatCount - 1).TestCase_index.Exists(DicKey_TestCase) Then
                                                SrcStock(ExistPatCount - 1).TestCase_index.Add DicKey_TestCase, CurTestNum
                                                
                                            Else
                                                theexec.Datalog.WriteComment "There are two same TestCase in Same Pattern"
                                                theexec.Datalog.WriteComment "Pattern Name :" & SrcStock(ExistPatCount - 1).PatternName
                                                theexec.Datalog.WriteComment "TestCase# :" & SrcStock(ExistPatCount - 1).TestCase_index(DicKey_TestCase) + 1 & "," & CurTestNum + 1 & vbCrLf
                                                GoTo errHandler
                                            End If
                                            If TempCount <> Len(SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStr) Then
                                                theexec.Datalog.WriteComment "Source Bit is NOT single bit" & vbCrLf
                                                theexec.Datalog.WriteComment "PatternName :" & SrcStock(ExistPatCount - 1).PatternName & vbCrLf
                                                theexec.Datalog.WriteComment "TestCase :" & SrcStock(ExistPatCount - 1).TestCase(CurTestNum).ConditionName & vbCrLf
                                            End If
                                            TempCount = 0
                                            Exit For
                                        End If
                                    Next TestCase_Row
                                    If CurColumn = MaxColumn Then
                                        CurTestNum = 0
                                        Flag_FindTestName = False
                                    End If
                                Else
                                    CurTestNum = 0
                                    Flag_FindTestName = False
                                End If
                            
                            Else
                                '''--------Find the Header "Test", record the row value and setup flag "Flag_FindTestName" as True--------
                                If UCase(vDigSrcInfo(PatternName_Row, CurColumn)) Like UCase("Test*") Then
                                    Flag_FindTestName = True
                                    '''--------Record the TestCase, i.g., Test1--------
                                    ReDim Preserve SrcStock(ExistPatCount - 1).TestCase(CurTestNum) As testCondition
                                        SrcStock(ExistPatCount - 1).TestCase(CurTestNum).ConditionName = vDigSrcInfo(PatternName_Row, CurColumn)
                                    '''--------Record the TestCase, i.g., Test1--------
                                    TempCount = 0
                                    For TestCase_Row = PatternName_Row + 1 To MaxRow
                                        tempStr = vbNullString
                                        If vDigSrcInfo(TestCase_Row, CurColumn) <> "" Then
                                            '''--------Record the TestCase, i.g., S or 0--------
                                            ReDim Preserve SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStrAry(TempCount) As String
                                            SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStrAry(TempCount) = vDigSrcInfo(TestCase_Row, CurColumn)
                                            '''--------Record the TestCase, i.g., S or 0--------
                                            TempCount = TempCount + 1
                                        Else
                                            '''--------Record the TestCase, i.g., SSSSSSSSS or 0000000000--------
                                            SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStr = Join(SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStrAry, "")
                                            '''--------Record the TestCase, i.g., SSSSSSSSS or 0000000000--------
                                            SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitCount = TempCount
                                            '''********add dictionary for index of all test cases ' swlinza 20190602********
                                            DicKey_TestCase = vbNullString
                                            DicKey_TestCase = SrcStock(ExistPatCount - 1).TestCase(CurTestNum).ConditionName
                                            SrcStock(ExistPatCount - 1).TestCase_index.compareMode = 1
                                            If Not SrcStock(ExistPatCount - 1).TestCase_index.Exists(DicKey_TestCase) Then
                                                SrcStock(ExistPatCount - 1).TestCase_index.Add DicKey_TestCase, CurTestNum
                                                
                                            Else
                                                theexec.Datalog.WriteComment "There are two same TestCase in Same Pattern"
                                                theexec.Datalog.WriteComment "Pattern Name :" & SrcStock(ExistPatCount - 1).PatternName
                                                theexec.Datalog.WriteComment "TestCase# :" & SrcStock(ExistPatCount - 1).TestCase_index(DicKey_TestCase) + 1 & "," & CurTestNum + 1 & vbCrLf
                                                GoTo errHandler
                                            End If
                                            If TempCount <> Len(SrcStock(ExistPatCount - 1).TestCase(CurTestNum).DigSrc_BitStr) Then
                                                theexec.Datalog.WriteComment "Source Bit is NOT single bit" & vbCrLf
                                                theexec.Datalog.WriteComment "PatternName :" & SrcStock(ExistPatCount - 1).PatternName & vbCrLf
                                                theexec.Datalog.WriteComment "TestCase :" & SrcStock(ExistPatCount - 1).TestCase(CurTestNum).ConditionName & vbCrLf
                                            End If
                                            TempCount = 0
                                            Exit For
                                        End If
                                    Next TestCase_Row
                                End If
                            End If
                            
                        Next CurColumn
                        
                        '''********add dictionary for index of pattern in SrcStock array swlinza 20190602********
                        DicKey_SrcStock = vbNullString
                        DicKey_SrcStock = vDigSrcInfo(CurRow, 1)
                        If Not Dic_SrcStockIndex.Exists(DicKey_SrcStock) Then
                            Dic_SrcStockIndex.Add DicKey_SrcStock, ExistPatCount - 1
                        Else
                            theexec.Datalog.WriteComment "There are two same patterns in control table"
                            theexec.Datalog.WriteComment "Pattern Name :" & DicKey_SrcStock
                            theexec.Datalog.WriteComment "Pattern# :" & Dic_SrcStockIndex.item(DicKey_SrcStock) & "," & ExistPatCount - 1 & vbCrLf
                            Debug.Print "There are two same patterns in control table"
                            GoTo errHandler
                        End If
                        
                    Else
                    '''--------If TempAry_SheetContent is empty, setup flag "Flag_FindPatternName" as False--------
                        Flag_FindPatternName = False
                    End If
                    
                Else
                '''--------Find the Header "Pattern Name", record the row value and setup flag "Flag_FindPatternName" as True--------
                    If vDigSrcInfo(CurRow, 1) Like "Pattern Name" Then
                        PatternName_Row = CurRow
                        Flag_FindPatternName = True
                    End If
                End If
            Next CurRow
            Application.ScreenUpdating = True
            DSSCMappingTableIsRead = True
        Else
            theexec.Datalog.WriteComment "There is no 'DSSCMappingTable' Sheet in this workbook"
        End If
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Parse_DSSCPat_DigSrcInfo") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''20210504 Get the system register data from OI.
Function RegKeyRead_autoZ(i_RegKey As String) As String
    Dim myWS As Object
    On Error Resume Next
    Set myWS = CreateObject("WScript.Shell")
    RegKeyRead_autoZ = myWS.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Teradyne\IG-XL\Handler Drivers\TELP8\" & i_RegKey)
End Function


''20210819 print OS information in datalog
' [20230731][T-Pal][HS] Parsing version info 1 time
Public Sub OS_Info()
On Error GoTo errHandler
    If glb_isOS_Info = False Then
    
        gls_OperatingSystem = CStr(TheHdw.Computer.OperatingSystem)
        gls_NumberOfProcessors = CStr(TheHdw.Computer.NumberOfProcessors)
        gls_ProcessingBits = CStr(TheHdw.Computer.ProcessingBits)
        gls_PhysicalMemory = CStr(TheHdw.Computer.PhysicalMemory)
        glb_isOS_Info = True
    End If

    With theexec.Datalog
        .WriteComment "-----------------------------------------------"
        .WriteComment "             OS Information"
        .WriteComment "-----------------------------------------------"
        .WriteComment "OS          : " & gls_OperatingSystem     'service pack
        .WriteComment "CPUnumbers  : " & gls_NumberOfProcessors  'number of processing cores
        .WriteComment "ProcessBits : " & gls_ProcessingBits
        .WriteComment "Memory      : " & gls_PhysicalMemory      'size of the main RAM
        .WriteComment vbNullString
    End With
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "OS_Info")
    If AbortTest Then Exit Sub Else Resume Next

End Sub


''20210819 print IGXL information in datalog
' [20230731][T-Pal][HS] Parsing version info 1 time
Public Sub IGXL_Info()
On Error GoTo errHandler
    If glb_isIGXL_Info = False Then
        gls_SoftwareVersion = CStr(theexec.SoftwareVersion)
        gls_SoftwareBuild = CStr(theexec.SoftwareBuild)
        glb_isIGXL_Info = True
    End If
    
    With theexec.Datalog
        .WriteComment "-----------------------------------------------"
        .WriteComment "             IGXL Information"
        .WriteComment "-----------------------------------------------"
        .WriteComment "IGXLVersion  : " & gls_SoftwareVersion  '
        .WriteComment "IGXLBuild    : " & gls_SoftwareBuild
        .WriteComment vbNullString
    End With
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "IGXL_Info")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


''20210819 print RunOption and EnableWrod information in datalog
Public Sub RunOptions_Info()
On Error GoTo errHandler
    Dim mAS_AllEnableWords() As String
    Dim mL_AllEnWrdsCnt As Long
    Dim mL_Loop As Long
    Dim mS_CFG_EW As String
 
    If glb_isParsing_EW = False Then
        mL_AllEnWrdsCnt = tl_ExecGetEnableWords(mAS_AllEnableWords)

        For mL_Loop = 0 To UBound(mAS_AllEnableWords)
            If theexec.enableWord(mAS_AllEnableWords(mL_Loop)) Then
                gls_Active_EW = CombineStringList(gls_Active_EW, mAS_AllEnableWords(mL_Loop), "|")
            Else
                gls_Disable_EW = CombineStringList(gls_Disable_EW, mAS_AllEnableWords(mL_Loop), "|")
            End If
        Next
        If isDebugMode = False Then
            glb_isParsing_EW = True
        End If
    End If
    
    If glb_isRunOptions_Info = False Then
        gls_DoAll = CStr(theexec.RunOptions.DoAll)
        gls_OverrideFailStop = CStr(theexec.RunOptions.OverrideFailStop)
        gls_Assume = CStr(theexec.RunOptions.Assume)
        gls_AssumeSiteDisable = CStr(theexec.Flow.RunOption(optAssumeSiteDisable))
        glb_isRunOptions_Info = True
    End If
    
    With theexec.Datalog
        .WriteComment "-----------------------------------------------"
        .WriteComment "          Run Options Information"
        .WriteComment "-----------------------------------------------"
        .WriteComment "DoAll                : " & gls_DoAll
        .WriteComment "OverrideFailStop     : " & gls_OverrideFailStop
        .WriteComment "Assume               : " & gls_Assume
        .WriteComment "AssumeSiteDisable    : " & gls_AssumeSiteDisable
        .WriteComment "Active EnableWords   : " & gls_Active_EW
        .WriteComment vbNullString
    End With
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "RunOptions_Info")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


''20210819 print sheet version in datalog
''1. PatSets_all sheet info
''2. Flow Char sheet
''3. Efuse DataBase version
''4. Vdd Binning version
''5. DC spec voltage version
''6. SPIROM sheet version
'[20230407][T-All] Add print info in PatSetAll sheet if header has key word "ver:"
' [20230731][T-Pal][HS] Parsing version info 1 time
Public Sub Version_Info()

    On Error GoTo errHandler
    
    Dim sFuncName As String:: sFuncName = "Version_Info"
    
    Dim wb As Workbook
    Dim ws_pat As Worksheet
    Dim ws_sheet As Worksheet
    
    Dim s_TestPlan_Key As String: s_TestPlan_Key = "VoltageTable"
    
    Dim s_sheetName As String
    'Dim s_PatVersion As String
    'Dim s_SCGHVersion As String
    Dim s_EfuseVersion As String
    Dim s_BinCutVersion As String
    'Dim s_TestPlanVersion As String
    Dim s_RTOSBinaryVersion As String
    Dim s_BinCutStarEQVersion As String
    'Dim s_PowerSettingVersion As String
    
    Dim s_TempAry() As String
    
    Dim i As Long
    Dim l_maxcol As Long
    Dim l_maxRow As Long
    Dim l_idxstamp As Long
    
    Dim b_isSheetFound As Boolean
        
    If glb_isVersion_Info = False Then
    
''    Call RTOS_Parse_Info
        ''20210819 print sheet version in datalog
        ''1. PatSets_all sheet info
        ''2. Test Plan info
        ''3. SCGH info
        ''4. DC spec voltage version
        ''5. Efuse DataBase version
        ''6. Vdd Binning version
        ''7. DC spec voltage version
        ''8. SPIROM sheet version
        ''9. Flow Char sheet

        Call Parse_PatSetAll
        Application.ScreenUpdating = False
        ''4. DC spec voltage version
        s_sheetName = "DC_Specs_" & UCase(currentJobName)
''    s_sheetName = "DC_Specs_CP1"
        Set wb = Application.ActiveWorkbook
        Set ws_sheet = wb.Sheets(s_sheetName)
        Call check_Sheet_Range(s_sheetName, wb, ws_sheet, l_maxRow, l_maxcol, b_isSheetFound)
        
        If b_isSheetFound = True Then
            Dim DC_Table_Ver As String
            Dim powersheet() As Variant
            Dim Index_Row As Long
            Dim Index_Col As Long
            powersheet = ws_sheet.range(Cells(1, 1), Cells(l_maxRow, l_maxcol)).value
            For Index_Col = 1 To l_maxcol
                If LCase(powersheet(4, Index_Col)) Like LCase("Comment*") Then
                    l_idxstamp = Index_Col
                    Exit For
                End If
            Next Index_Col
            For Index_Row = 1 To l_maxRow
                If LCase(powersheet(Index_Row, l_idxstamp)) Like LCase("VoltageTable*") Or _
                    LCase(powersheet(Index_Row, l_idxstamp)) Like LCase("TestSetting*") Then '''//Revision of Voltage tables
                    gls_DC_Table_Ver = powersheet(Index_Row, l_idxstamp)
                    Exit For
                End If
            Next Index_Row
            Application.ScreenUpdating = True
        End If
    

        If gls_TestPlanVersion = "" Then gls_TestPlanVersion = "N/A"
        If gls_PatVersion = "" Then gls_PatVersion = "N/A"
        If gls_SCGHVersion = "" Then gls_SCGHVersion = "N/A"
        If gls_DC_Table_Ver = "" Then gls_DC_Table_Ver = "N/A"
        
        glb_isVersion_Info = True
    End If
    'TheExec.Datalog.WriteComment "*******************program information*******************"
    With theexec.Datalog
        .WriteComment "-----------------------------------------------"
        .WriteComment "          Sheet and Version Infomation"
        .WriteComment "-----------------------------------------------"
        .WriteComment "print: Current job name          :" & UCase(currentJobName)
        .WriteComment "print: Test plan version         :" & gls_TestPlanVersion
        .WriteComment "print: Pattern list version      :" & gls_PatVersion
        .WriteComment "print: SCGH version              :" & gls_SCGHVersion
        .WriteComment "print: DC_spec PowerSetting      :" & gls_DC_Table_Ver
''        .WriteComment "print: Char plan version         :" & gS_Char_Version
''        .WriteComment "print: Efuse DataBaseVerision    :" & EfuseVersion
''        .WriteComment "print: BinCut Version            :" & BinCutVersion
''        .WriteComment "print: RTOS version              :" & binary_ver
        .WriteComment vbNullString

        If UBound(gl_saPatSetAllPrintInfo) > -1 Then
            For i = 0 To UBound(gl_saPatSetAllPrintInfo)
                If gl_saPatSetAllPrintInfo(i) <> "" Then
                    .WriteComment gl_saPatSetAllPrintInfo(i)
                End If
            Next i
        End If
    End With

Exit Sub
                                                                                                                                                                                                                                                              
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, sFuncName)
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Function TDR_ExcludedPin_WriteTraceLen(excludePins As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim TraceLen As Double
    Dim chArray() As String
    Dim chArray_1() As String
    Dim siteNum As Long
    Dim retPinNames As String
    Dim retSiteNums() As Long
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim z As Integer
    Dim exPins() As String
    Dim excludeNumberPins As Long
    Dim typesCount As Long
    Dim stringTypes() As String

    siteNum = theexec.sites.Existing.Count
    Call theexec.DataManager.DecomposePinList(excludePins, exPins(), excludeNumberPins)  ''pins do not need calibrate
    ReDim chArray(excludeNumberPins - 1) As String
    ReDim chArray_1((siteNum * excludeNumberPins) - 1) As String
    
    j = 0
    theexec.DataManager.ReturnSignalNames = True    ''true = channel mode, Acture Waveform can not work
    For Each site In theexec.sites
        For i = 0 To excludeNumberPins - 1
            Call theexec.DataManager.GetChannelStringFromPinAndSite(exPins(i), site, chArray(i))
            Call theexec.DataManager.GetChannelTypes(exPins(i), typesCount, stringTypes())  ''get channels from input pingroup
            If UCase(stringTypes(0)) = "N/C" Then chArray(i) = vbNullString    ''if pin type N/C, then ignore channel assignment, and with empty
            chArray_1(j) = chArray(i)
            j = j + 1
        Next i
    Next site

    If glb_TesterType = "Jaguar" Then
        For i = 0 To ((siteNum * excludeNumberPins) - 1)
            If chArray_1(i) <> "" Then
                TheHdw.Digital.Calibration.Channels(chArray_1(i)).DIB.trace = 0.000000001
            End If
        Next i
    ElseIf glb_TesterType = "UltraFLEXplus" Then
        For i = 0 To ((siteNum * excludeNumberPins) - 1)
            If chArray_1(i) <> "" Then
''                thehdw.Digital.Calibration.Channels(chArray_1(i)).trace = 0.000000001  ''API commnad only for UFP
                TheHdw.Calibration.DIB.Channels(chArray_1(i)).trace = 0.000000001
            End If
        Next i
    End If
    theexec.DataManager.ReturnSignalNames = False  ''false = pogo mode
    
    If glb_TesterType = "UltraFLEXplus" Then
        TheHdw.Calibration.DIB.Traces.Apply
    End If

    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "TDR_ExcludedPin_WriteTraceLen") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Print_IProfileValue(PinName As String, TestName As String, CurrentProfile As DSPWave, site As Variant, Optional percent_control As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim TestNumber As Long
    Dim max As Double
    Dim min As Double
    Dim avg As Double
    Dim latestTenMean As Double
    Dim percent As Double
    Dim dspw As New DSPWave
    Dim T1 As Double
    Dim T2 As Double
    Dim CurrentProfile_Sort As New DSPWave
    Dim CurrentProfile_SUM As New DSPWave
    Dim CurrentProfile_Ten As New DSPWave
    
    If percent_control = 0 Then percent_control = 1
    
    TestNumber = theexec.sites.item(site).TestNumber
    ''=============offline debug=====================
    If theexec.TesterMode = testModeOffline Then
       dspw.CreateRandom 0, 100, 1000
       CurrentProfile = dspw.Copy
    End If
    ''===============================================
    CurrentProfile_Sort = CurrentProfile.Sort
    CurrentProfile_SUM = CurrentProfile_Sort.Select(0, 1, (UBound(CurrentProfile_Sort.data) + 1) * percent_control)
    percent = CurrentProfile_SUM.CalcSum / (CurrentProfile.sampleSize * percent_control)
    
    latestTenMean = CurrentProfile_Ten.Select(CurrentProfile.sampleSize * (1 - 0.1), , CurrentProfile_Ten.sampleSize).CalcMean
    avg = CurrentProfile_Sort.CalcMean
    max = CurrentProfile_Sort.CalcMaximumValue
    min = CurrentProfile_Sort.CalcMinimumValue
    
    If max >= 0 Then
        theexec.Flow.TestLimit resultVal:=max, Tname:=TestName, PinName:=PinName & "_" & "current_maximum", ForceResults:=tlForceNone, lowVal:=min, hiVal:=max * 1.01
    Else
        theexec.Flow.TestLimit resultVal:=max, Tname:=TestName, PinName:=PinName & "_" & "current_maximum", ForceResults:=tlForceNone, lowVal:=min, hiVal:=max * 0.99
    End If
    theexec.Flow.TestLimit resultVal:=percent, Tname:=TestName, PinName:=PinName & "_" & "current_" & percent_control * 100 & "percent", ForceResults:=tlForceNone, lowVal:=min, hiVal:=max
    theexec.Flow.TestLimit resultVal:=avg, Tname:=TestName, PinName:=PinName & "_" & "current_average", ForceResults:=tlForceNone, lowVal:=min, hiVal:=max
    If min >= 0 Then
        theexec.Flow.TestLimit resultVal:=min, Tname:=TestName, PinName:=PinName & "_" & "current_minimum", ForceResults:=tlForceNone, lowVal:=min * 0.99, hiVal:=max
    Else
        theexec.Flow.TestLimit resultVal:=min, Tname:=TestName, PinName:=PinName & "_" & "current_minimum", ForceResults:=tlForceNone, lowVal:=min * 1.01, hiVal:=max
    End If
    theexec.Flow.TestLimit latestTenMean, Tname:=TestName, PinName:=PinName & "_" & "current_latestTenMean", ForceResults:=tlForceNone
    '' 20160218 - Modify sequence to let TestResult after WriteFunctionalResult to cover test number increment 2 issue if souce sink time out alarm happen.
    theexec.sites.item(site).testResult = siteFail
    BV_Pass = False
    theexec.sites.item(site).TestNumber = TestNumber + 1
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "Print_IProfileValue") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function


'20211220: Added enablePrint to control the error datalog print
'20200703: Created to check range of row and column for the sheet.
Public Function check_Sheet_Range(sheetName As String, wb As Workbook, ws_def As Worksheet, MaxRow As Long, maxcol As Long, isSheetFound As Boolean, Optional enablePrint As Boolean = True)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    If Find_Sheet(sheetName) = True Then
        wb.Sheets(sheetName).Unprotect
        Set ws_def = wb.Sheets(sheetName)
        ws_def.Select
    
        '''//Check ranges of row and column
        MaxRow = ws_def.Cells.SpecialCells(xlCellTypeLastCell).Row
        maxcol = ws_def.Cells.SpecialCells(xlCellTypeLastCell).Column
        
        If MaxRow > 0 And maxcol > 0 Then
            isSheetFound = True
        Else
            isSheetFound = False
            MaxRow = 0
            maxcol = 0
            If enablePrint = True Then
                theexec.Datalog.WriteComment "Content of " & sheetName & " is empty or incorrect. Error!!!"
                theexec.ErrorLogMessage "Content of " & sheetName & " is empty or incorrect. Error!!!"
            End If
        End If
    Else
        isSheetFound = False
        MaxRow = 0
        maxcol = 0
        If enablePrint = True Then
            theexec.Datalog.WriteComment sheetName & " doesn't exist in this workbook. Error!!!"
            theexec.ErrorLogMessage sheetName & " doesn't exist in this workbook. Error!!!"
        End If
    End If
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "check_Sheet_Range") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Sub PrintIFoldInfo(list As String)      'print suggest iFold limit
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If theexec.Flow.enableWord("PrintIFoldInfo") Then
    
        Dim power_pin_ary() As String
        Dim nNumOfPins As Long
        Dim sInstanceName As String
        sInstanceName = LCase(theexec.DataManager.instancename)
        
        If (Not sInstanceName Like "evs*") And (Not sInstanceName Like "gfxmbist*") And (Not sInstanceName Like "socmbist*") _
                            And (Not sInstanceName Like "cpumbist*") And (Not sInstanceName Like "m**0*_*") Then
            Exit Sub
        End If
        
        Call theexec.DataManager.DecomposePinList(list, power_pin_ary, nNumOfPins)
    
        Dim v_site As Variant
        Dim pin As Variant
        Dim i As Integer
        Dim sIFoldInfo As String
        Dim sdCurrentRange As New SiteDouble
        Dim sdFoldLimit As New SiteDouble
        Dim sdSuggestIFoldLimit As New SiteDouble
        
        For Each pin In power_pin_ary
            sdSuggestIFoldLimit = 0
            sdCurrentRange = TheHdw.DCVS.Pins(pin).Meter.CurrentRange.value
            sdFoldLimit = TheHdw.DCVS.Pins(pin).CurrentLimit.Source.FoldLimit.level.value
            sdSuggestIFoldLimit = sdSuggestIFoldLimit.Add(sdFoldLimit).Subtract(sdCurrentRange.Multiply(0.125))
            
            For Each v_site In theexec.sites.Active
                
                sIFoldInfo = sIFoldInfo + "IFoldInfo: " + sInstanceName + "," + _
                                pin + "," + _
                                CStr(sdCurrentRange(v_site)) + "," + _
                                CStr(sdFoldLimit(v_site)) + "," + _
                                CStr(sdSuggestIFoldLimit(v_site))
                
                theexec.Datalog.WriteComment sIFoldInfo
                sIFoldInfo = vbNullString
                
            Next
        Next
    End If
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "PrintIFoldInfo") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


'[20230407][T-All] Add nCommentColumn position string to store in gl_saPatSetAllPrintInfo
'[20230511][T-All] Modify to fix ver print two times.
Public Function Parse_PatSetAll()
On Error GoTo errHandler
    Dim sFuncName As String:: sFuncName = "Parse_PatSetAll"

    Dim MySheet As Worksheet:: Set MySheet = Nothing
    Dim sSheetName As String
    Dim nMaxColumn As Long, nMaxRow As Long
    Dim nCurRow As Long
    Dim vPatSetAllArr As Variant
    Dim s_TempAry() As String
    Dim s_TempComment As String

    Dim nStartRow As Integer: nStartRow = 4
    Dim nPatternSetColumn As Integer: nPatternSetColumn = 2
    Dim nFileGroupNameColumn As Integer: nFileGroupNameColumn = 5
    Dim nCommentColumn As Integer: nCommentColumn = 9
    Dim s_ErrorStatement As String
    Dim i As Integer: i = 0
    
    If gl_isParsePatSetAll = False Then
    
        gl_isCheckPreConditionVerDone = False     ' inital need to check PreConditionVersion
        
        PatSetAllDic.RemoveAll
        ReDim gl_saPatSetAllPrintInfo(0)
    '==============================================================================
    'Parsing PatSets_All Sheet data.
    '==============================================================================
        sSheetName = "PatSets_All"
        If WorksheetExists(sSheetName, False) Then
            Application.ScreenUpdating = False
            Set MySheet = ActiveWorkbook.Sheets(sSheetName)
            MySheet.Select
            nMaxRow = MySheet.UsedRange.Rows.Count
            nMaxColumn = MySheet.UsedRange.Columns.Count
            vPatSetAllArr = MySheet.range(Cells(nStartRow, nPatternSetColumn), Cells(nMaxRow, nCommentColumn)).value
            For nCurRow = 1 To UBound(vPatSetAllArr)
                If vPatSetAllArr(nCurRow, nPatternSetColumn - 1) = "" Then
                    Exit For
                Else
                    If Not PatSetAllDic.Exists(UCase(Trim(vPatSetAllArr(nCurRow, nPatternSetColumn - 1)))) Then
                        PatSetAllDic.Add UCase(Trim(vPatSetAllArr(nCurRow, nPatternSetColumn - 1))), UCase(Trim(vPatSetAllArr(nCurRow, nFileGroupNameColumn - 1)))
                    End If
                    
                    If vPatSetAllArr(nCurRow, nCommentColumn - 1) Like "ver:*" Then
                        s_TempComment = Replace(vPatSetAllArr(nCurRow, nCommentColumn - 1), "ver:", "", 1, 1)
                        If UCase(s_TempComment) Like UCase("*pattern*dashboard*:*") Then
                            s_TempAry = Split(s_TempComment, ":")
                            gls_PatVersion = Trim(s_TempAry(1))
                        ElseIf UCase(s_TempComment) Like UCase("*TestPlan*:*") Then
                            s_TempAry = Split(s_TempComment, ":")
                            gls_TestPlanVersion = Trim(s_TempAry(1))
                        ElseIf UCase(s_TempComment) Like UCase("*SCGH*:*") Then
                            s_TempAry = Split(s_TempComment, ":")
                            gls_SCGHVersion = Trim(s_TempAry(1))
                        Else
                            gl_saPatSetAllPrintInfo(i) = Trim(s_TempComment)
                            i = i + 1
                            ReDim Preserve gl_saPatSetAllPrintInfo(i)
                        End If
                    End If
                End If
            Next nCurRow
            Set MySheet = Nothing
            Application.ScreenUpdating = True
        Else
            s_ErrorStatement = sSheetName & " doesn't exist in current program. Error!!!"
            Call Print_Error_Message(Warning_Info, LIB_Common, sFuncName, s_ErrorStatement)      '20221205 Tank try use print error
            Exit Function
        End If

        gl_isParsePatSetAll = True
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, sFuncName)      '20221205 Tank try use print error
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function Check_PreCondition_Pattern_Version()

On Error GoTo errHandler

    Dim sFuncName As String:: sFuncName = "Check_PreCondition_Pattern_Version"
    Dim n As Long
    Dim TempArray() As String
    Dim s_ErrorStatement As String
    
    If gl_isCheckPreConditionVerDone = False Then
        If gl_isParsePreConditionDone And gl_isParsePatSetAll Then   'check Parse PreCondition sheet and PatSerAll sheet
            If Compare_Dictionary_Exist(PatSetAllDic) And gl_ePreConditionError = PreConditionError.PreConditionPass Then     'check PreCondition sheet or PatSetAll sheet data
                For n = 0 To UBound(PreConditionInfo)
                    If PatSetAllDic.Exists(UCase(PreConditionInfo(n).PatternName)) Then     'check PreCondition sheet pattern in PatSetAll sheet
                        TempArray = Split(PatSetAllDic.item(UCase(PreConditionInfo(n).PatternName)), ":")
                        If UBound(TempArray) = 1 Then
                            If TempArray(1) <> UCase(PreConditionInfo(n).PatternVersion) Then       'check version
                                gl_ePreConditionError = PreConditionError.PatVerDifferent   'bIsMatch = False
                                Exit For
                            End If
                        Else
                            gl_ePreConditionError = PreConditionError.PatVerDifferent   'bIsMatch = False
                            Exit For
                        End If
                    Else
                        gl_ePreConditionError = PreConditionError.PatVerDifferent   'bIsMatch = False
                        Exit For
                    End If
                Next n
            Else
                gl_ePreConditionError = PreConditionError.PatVerDifferent   'bIsMatch = False
            End If
        Else
            gl_ePreConditionError = PreConditionError.ParseSheetError
        End If
        
        gl_isCheckPreConditionVerDone = True
        
    End If
    
    PatSetAllDic.RemoveAll  '20230406 no need use in other function
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, sFuncName)      '20221205 Tank try use print error
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Compare_PreConditionInfo_Exist(TempArray() As preCondition) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Compare_PreConditionInfo_Exist = False
    If UBound(TempArray) >= 0 Then
        Compare_PreConditionInfo_Exist = True
    End If
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    err.Clear
    Compare_PreConditionInfo_Exist = False
End Function


Public Function Compare_Dictionary_Exist(TempDic As Dictionary) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Compare_Dictionary_Exist = False
    If TempDic.Count <> 0 Then
        Compare_Dictionary_Exist = True
    End If
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    err.Clear
    Compare_Dictionary_Exist = False
End Function


' [20230801][All][Tank] modify function to avoid memory leak
Public Function CombineStringList(sOrgString As String, sAddNewString As String, Optional sSymbol As String = ",", Optional bIgnoreEmpty As Boolean = True) As String   'ex: pin1,pin2,pin3,pin4
On Error GoTo errHandler
    Dim sFuncName As String:: sFuncName = "CombineStringList"
    Dim sa_tmpStr() As String
    ReDim sa_tmpStr(1)
    sa_tmpStr(0) = sOrgString
    sa_tmpStr(1) = sAddNewString
    
    If sa_tmpStr(0) = "" Then
        CombineStringList = sAddNewString
    Else
        If bIgnoreEmpty And sa_tmpStr(1) = "" Then
            CombineStringList = sa_tmpStr(0)
        Else
            CombineStringList = Join(sa_tmpStr, sSymbol)
        End If
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, sFuncName)
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Print_Error_Message(ByVal Error_Warning As Error_Warning_Info, ByVal s_ModuleName As String, ByVal s_FuncName As String, Optional s_Statement As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim s_temp As String
    Dim s_TempMsg As String
    
    s_temp = LCase(s_ModuleName)
    If s_temp Like "*vdd*binning*" Then
        glb_SVN = VDDBINNING_VER
    ElseIf s_temp Like "*common*" Then
        glb_SVN = COMMON_VER
    ElseIf s_temp Like "*dc*" Or s_temp Like "*evs*" Then
        glb_SVN = DC_VER
    ElseIf s_temp Like "*digital*functional*t*" Or s_temp Like "*harvest*" Then
        glb_SVN = DIGITAL_VER
    ElseIf s_temp Like "*mbist*" Then
        glb_SVN = MBIST_VER
    ElseIf s_temp Like "*shmoo*" Or s_temp Like "*digital*debug" Then
        glb_SVN = CZ_VER
    ElseIf s_temp Like "*efuse*" Then
        glb_SVN = EFUSE_VER
    ElseIf s_temp Like "*hardip*" Or s_temp Like "*dactrim*" Or s_temp Like "*k*addrio*" Then
        glb_SVN = HIP_VER
    ElseIf s_temp Like "*rtos*" Or s_temp Like "*uart*" Then
        glb_SVN = SPIROM_VER
    
    ElseIf s_temp Like "*otp*" Or s_temp Like "*ahb*" Then
        glb_SVN = LCD_OTP_VER
    ElseIf s_temp Like "*dvdc*" Or s_temp Like "*witrimming*" Or s_temp Like "*lib*tname*" Or s_temp Like "*dssc*setup*" Or s_temp Like "*lib*onchip*" Or s_temp Like "*spotcal*" Then
        glb_SVN = LCD_DVDC_VER
    Else
        glb_SVN = "9999.9999.9999"
    End If
    
    s_TempMsg = IIf(s_Statement = "", ":: Please check it out.", s_Statement)
    If Error_Warning = Error_Info Then
        theexec.Datalog.WriteComment "!!!! Error Info: " + glb_SVN + ", " + s_ModuleName + ", " + s_FuncName + ", " + s_TempMsg + " !!!!"
        'TheExec.Datalog.WriteComment "<Error>: " + s_TempMsg
    ElseIf Error_Warning = Warning_Info Then
        theexec.Datalog.WriteComment "!!!! Warning Info: " + glb_SVN + ", " + s_ModuleName + ", " + s_FuncName + ", " + s_TempMsg + " !!!!"
        'TheExec.Datalog.WriteComment "<Warning>: " + s_TempMsg
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "Print_Error_Message") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230424][All][Tank] add init Flag_GetCurrentProfile flag
' [20230628][All][Gillian] add Initial HardIP flag
' [20231108][All][Tank] Add ParseIDSMappingTable in OnProgramStarted first
Public Sub InitVariableOnValidated()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Flag_RAK_INIT = False   'Not initialized, Init in HardIP flow
    Flag_RSCR_INIT = False  'Not initialized, Init in MBist flow
    Flag_Shmoo_INIT = False 'Not initialized, Init in Shmoo flow
    Flag_MBISTFailBlock_INIT = False
    Flag_GetChannelType = False
    Flag_GetPowerSeq = False
    Flag_DC_ParsingRecoveryTable = False  '20211008 add for settle time
    Flag_GetCurrentProfile = False

    gl_isCheckPreConditionVerDone = False
    gl_isParsePreConditionDone = False
    gl_isParsePatSetAll = False
    gl_isCheckClampLimit = ContiClampCheckType.CheckInit    'Check DC limit on first TD. Check function => CheckTestInst_HiLoLimit()'
    gl_PreCondition_Dic.RemoveAll
    PatSetAllDic.RemoveAll
        
    gl_isPatAndCZmappingTable = False
    gl_isFind_nWire_Pin = False
        
    gl_GetInstrument_Dic.RemoveAll
    gl_GetInstrumentType_Dic.RemoveAll
    gl_dicPowerPinIndex.RemoveAll

    gl_isParExecutionProfileSheet = False
    gl_EnableCurrentProfile = False
    gl_EnableVoltageProfile = False
    gl_nWireFreq = 0

    Erase PreConditionInfo
        
    gl_nWireFreq_Value_Dict.RemoveAll
    gl_nWireFreq_AC_Dict.RemoveAll
        
    EnableDigitalTestLimitTTR = False
    EnableFieldProcesingTTR = False
    gl_Disable_HIP_debug_log = False
    glb_Disable_CurrRangeSetting_Print = False
    EnableHardIPDigCapsdisableTTR = False
    EnableAnalogMuxOutTTR = False
    EnableHardIPTnameConstructionTTR = False
    GLB_InstanceParaDict.RemoveAll
    
    glb_isVersion_Info = False
    glb_isOS_Info = False
    glb_isIGXL_Info = False
    glb_isRunOptions_Info = False
    glb_isParsing_EW = False

    gls_PatVersion = vbNullString
    gls_TestPlanVersion = vbNullString
    gls_DC_Table_Ver = vbNullString

    gls_OperatingSystem = vbNullString
    gls_NumberOfProcessors = vbNullString
    gls_ProcessingBits = vbNullString
    gls_PhysicalMemory = vbNullString

    gls_SoftwareVersion = vbNullString
    gls_SoftwareBuild = vbNullString

    gls_DoAll = vbNullString
    gls_OverrideFailStop = vbNullString
    gls_Assume = vbNullString
    gls_AssumeSiteDisable = vbNullString
    gls_Active_EW = vbNullString
    gls_Disable_EW = vbNullString

    Flag_DTS_function = False
    Flag_BurstPat_INIT = False
    DSSCMappingTableIsRead = False
    glb_FlagCheckingFRCClock = False '20230524 add for FRC
    Flag_IDSMappingTable = False
    glb_CheckIDSMappingTable_With_Fuse = False
    glb_isParsingSFC = False
    SkipStep1_Flag = False
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "InitVariableOnValidated") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


' [20231003][All][Tank] modify after Chihome review
Public Sub SetPowerAndIOPin_Volt_0v(sPowerPinName As String, sDigitalPinName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim sFuncName As String:: sFuncName = "SetPowerAndIOPin_Volt_0v"

    Dim pins_power() As String
    Dim nPowerPinCnt As Long
    Dim vPowerPin As Variant
    Dim pins_digital() As String
    Dim nDigitalPinCnt As Long
    Dim vDigitalPin As Variant
    Dim sPinName As String
    Dim sSlotType As String
    Dim sOtherPin As String
    Dim sTempDCVS_pin As String
    Dim sTempDCVI_pin As String
    
    
    If sPowerPinName <> "" Then
        theexec.DataManager.DecomposePinList sPowerPinName, pins_power(), nPowerPinCnt
        If nPowerPinCnt > 0 Then
            For Each vPowerPin In pins_power    '20230406 modify digital pin set 0v parallel
                If gl_GetInstrumentType_Dic.Exists(LCase(sPinName)) Then
                    sPinName = vPowerPin
                    If LCase(gl_GetInstrumentType_Dic(LCase(sPinName))) Like "*dcvs*" Then
                        sTempDCVS_pin = CombineStringList(sTempDCVS_pin, sPinName)
                    ElseIf LCase(gl_GetInstrumentType_Dic(LCase(sPinName))) Like "*dcvi*" Then
                        sTempDCVI_pin = CombineStringList(sTempDCVI_pin, sPinName)
                    Else
                        sOtherPin = CombineStringList(sOtherPin, sPinName)
                    End If
                Else
                    sOtherPin = CombineStringList(sOtherPin, sPinName)
                End If
            Next vPowerPin
                
            If sTempDCVS_pin <> "" Then
                With TheHdw.DCVS.Pins(sTempDCVS_pin)
                    .Voltage.Main.value = 0
                    .Connect
                    .Gate = tlOn
                    TheHdw.Wait 0.01
                    .Gate = tlOff
                    .Disconnect
                End With
            Else
            End If
            
            If sTempDCVI_pin <> "" Then
                With TheHdw.DCVI.Pins(sTempDCVI_pin)
                    .Voltage = 0
                    .Connect
                    .Gate = tlOn
                    TheHdw.Wait 0.01
                    .Gate = tlOff
                    .Disconnect
                End With
            Else
            End If
            
            If sOtherPin <> "" Then
                theexec.ErrorLogMessage ">>>> Waring Info: " & COMMON_VER & ", " & "LIB_Common" & ", " & sFuncName & " <<<<"     'need to check power pin instrument
                theexec.ErrorLogMessage "<Waring> " & "PowerPin : " & sOtherPin & " is no define!!!"
            End If
            'TheExec.AddOutput "All power pins force 0V"
        Else
            'TheExec.ErrorLogMessage "PowerPin : " & sPowerPinName & " is no define!!!"
            theexec.ErrorLogMessage ">>>> Error Info: " & COMMON_VER & ", " & "LIB_Common" & ", " & sFuncName & " <<<<"     'need to modify sub function
            theexec.ErrorLogMessage "<Error> " & "PowerPin : " & sPowerPinName & " is no define!!!"
        End If
    End If
    
    If sDigitalPinName <> "" Then
        theexec.DataManager.DecomposePinList sDigitalPinName, pins_digital(), nDigitalPinCnt
        If nDigitalPinCnt > 0 Then      '20230406 modify digital pin set 0v parallel
            With TheHdw.PPMU.Pins(sDigitalPinName)
                .ForceV 0, 0.000002, 0.6, -0.6
                .ClampVHi = 6.5
                .ClampVLo = -1.5
                .Connect
                .Gate = tlOn
                TheHdw.Wait 0.01
                .Gate = tlOff
                .Disconnect
            End With
        Else
            'TheExec.ErrorLogMessage "DigitalPin : " & sDigitalPinName & " is no define!!!"
            theexec.ErrorLogMessage ">>>> Error Info: " & COMMON_VER & ", " & "LIB_Common" & ", " & sFuncName & " <<<<"     'need to modify sub function
            theexec.ErrorLogMessage "<Error> " & "DigitalPin : " & sDigitalPinName & " is no define!!!"
        End If

    End If
    
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "SetPowerAndIOPin_Volt_0v") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub



''20230223: Modidfied to add Parse Execution Profile Sheet, use by new current profile.
' [20230809][All][Oliver] modify for current profile other sheet method
Public Sub ParseExecutionProfileSheet()
    On Error GoTo errHandler
    Dim funName As String: funName = "ParseExecutionProfileSheet"
    
    Dim Excution_Profile() As Excution_Profile_Info
    
    'string
    Dim sSheetName As String
    Dim ParsingArr() As Variant
    'long
    Dim nSheetMaxRow As Long
    Dim nSheetMaxCol As Long
    
    Dim nStartRow As Long
    Dim dicFlowName As New Dictionary
    Dim nCurrentRowPos As Long
    Dim nCurrentColPos As Long
    Dim bJobMatch As Boolean
    Dim R As Long
    Dim n As Long
    Dim i As Long, j As Long
    Dim nTimeCol As Long
    Dim nFlowStartIndex As Long
    Dim nFlowEndIndex As Long
    Dim nTempCnt As Long
    Dim dTempTime As Double
    
'    Const nHeaderItemCol = 2
'    Dim nHeaderValueCol as long
'    Const nHeaderValueCol = 4
'    Const nFlowNameCol = 2
    Const nInstanceNameCol = 4
    Const nInstanceNameLine = 3
'    If LCase(glb_TesterType) = "jaguar" Then
'       nTimeCol = 5
'    Else
'       nTimeCol = 10
'    End If
    Dim nFlowNameCol As Long
    Dim nTimeRow As Long
    bJobMatch = False
    
    If gl_isParExecutionProfileSheet = False Then
        ''-------------------------
        ''Parse Execution Profile Sheet
        ''-------------------------
        sSheetName = "Execution Profile"
        If GetSheetInfo(sSheetName, nSheetMaxRow, nSheetMaxCol, ParsingArr) Then
            ReDim Excution_Profile(0)
            dicFlowName.RemoveAll
            glb_ExecutionProfile_Dict.RemoveAll
            
            If nSheetMaxRow > 0 And nSheetMaxCol > 0 Then
                For i = 1 To nSheetMaxRow
                    For j = 1 To nSheetMaxCol
                        If ParsingArr(i, j) = "Job" Then
                            nCurrentRowPos = i
                            nCurrentColPos = j
                            GoTo skip
                        End If
                        If nCurrentRowPos > 0 Then
                            If ParsingArr(i, j) <> "" And ParsingArr(i, j) = theexec.CurrentJob Then
                                bJobMatch = True
                                Exit For
                            ElseIf ParsingArr(i, j) <> "" Then
                                bJobMatch = False
                                Call Print_Error_Message(Error_Info, "LIB_Common", "ParseExecutionProfileSheet", "Error current job not match excution profile sheet! ")
                                Exit For
                            End If
                        End If

skip:
                    Next j
                    If bJobMatch Then Exit For
                Next i
                    
                If bJobMatch Then
                    'Find flow count
                    For R = nCurrentRowPos To nSheetMaxRow
                        If ParsingArr(R, nCurrentColPos) = "Flow Sheet" Then
                            nTimeRow = R + 1
                            nStartRow = R + 2
                            If nTimeRow > 0 Then
                                For j = 1 To nSheetMaxCol
                                    If ParsingArr(nTimeRow, j) = "Total" Then
                                        nTimeCol = j
                                        Exit For
                                    End If
                                Next j
                            End If
                            Exit For
                        End If
                    Next R
                    Dim nFlowCnt As Long
                    If nStartRow > 0 Then
                        For R = nStartRow To nSheetMaxRow
                            If dicFlowName.Exists(ParsingArr(R, nCurrentColPos)) = False And _
                                IgnoreFlow(CStr(ParsingArr(R, nCurrentColPos))) = False Then
                                dicFlowName.Add ParsingArr(R, nCurrentColPos), nFlowCnt
                                ReDim Preserve Excution_Profile(nFlowCnt)
                                Excution_Profile(nFlowCnt).sFlowName = ParsingArr(R, nCurrentColPos)
                                
                                nFlowCnt = nFlowCnt + 1
                            End If
                        Next R
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_Common", "ParseExecutionProfileSheet", _
                        "Error excution profile sheet no data! ")
                    Exit Sub
                    End If
                    
                    'Calculate flow(sub flow) process total time

                    
                    For n = 0 To (dicFlowName.Count - 1)
                        dTempTime = 0#
                        nFlowStartIndex = 0
                        nFlowEndIndex = 0
                        nTempCnt = 0
                        'find flow range
                        For R = nStartRow To nSheetMaxRow
                            If dicFlowName.Exists(ParsingArr(R, nCurrentColPos)) Then
                                If dicFlowName.item(ParsingArr(R, nCurrentColPos)) = n Then
                                    If nTempCnt = 0 Then
                                        nFlowStartIndex = R     'start pos
                                    End If
                                    nFlowEndIndex = R   'end pos
                                    nTempCnt = nTempCnt + 1     'check is flow first item
                                End If
                            End If
                        Next R
                        'calculate flow time between start pos to end pos
                        For R = nFlowStartIndex To nFlowEndIndex
                            dTempTime = dTempTime + IIf(IsNumeric(ParsingArr(R, nTimeCol)), CDbl(ParsingArr(R, nTimeCol)), 0#)
                        Next R
                        
                        'store to dict
                        Excution_Profile(n).dTotalProcessTime = dTempTime
                        If glb_ExecutionProfile_Dict.Exists(Excution_Profile(n).sFlowName) = False Then
                            glb_ExecutionProfile_Dict.Add LCase(Excution_Profile(n).sFlowName), _
                            Excution_Profile(n).dTotalProcessTime
                        End If
                    Next n
                    
                    'Input instance and instance time
                    For R = nStartRow To nSheetMaxRow
                        If IgnoreItem(CStr(ParsingArr(R, nInstanceNameCol))) = False Then
                            If InStr(LCase(CStr(ParsingArr(R, nCurrentColPos)) & "_" & CStr(ParsingArr(R, nInstanceNameCol))), "(characterize)") > 0 Then
                                ParsingArr(R, nInstanceNameCol) = Replace(ParsingArr(R, nInstanceNameCol), " (characterize)", "")
                            Else
                            End If
                            
                            If glb_ExecutionProfile_Dict.Exists(LCase(CStr(ParsingArr(R, nCurrentColPos)) & "_" & CStr(ParsingArr(R, nInstanceNameCol))) & "_" & CStr(ParsingArr(R, nInstanceNameLine))) = False Then
                                glb_ExecutionProfile_Dict.Add LCase(CStr(ParsingArr(R, nCurrentColPos)) & "_" & CStr(ParsingArr(R, nInstanceNameCol))) & "_" & CStr(ParsingArr(R, nInstanceNameLine)), CDbl(ParsingArr(R, nTimeCol))
                            Else
                                If glb_ExecutionProfile_Dict(LCase(CStr(ParsingArr(R, nCurrentColPos)) & "_" & CStr(ParsingArr(R, nInstanceNameCol)) & "_" & CStr(ParsingArr(R, nInstanceNameLine)))) < CDbl(ParsingArr(R, nTimeCol)) Then
                                   glb_ExecutionProfile_Dict(LCase(CStr(ParsingArr(R, nCurrentColPos)) & "_" & CStr(ParsingArr(R, nInstanceNameCol))) & "_" & CStr(ParsingArr(R, nInstanceNameLine))) = CDbl(ParsingArr(R, nTimeCol))
                                Else
                                End If
                            End If
                        End If
                    Next R
                    gl_isParExecutionProfileSheet = True
                Else
                
                End If
            Else
                Call Print_Error_Message(Error_Info, "LIB_Common", "ParseExecutionProfileSheet", _
                "Error excution profile sheet no data! ")
            End If
        Else
            Call Print_Error_Message(Error_Info, "LIB_Common", "ParseExecutionProfileSheet", _
        "Error excution profile sheet not exist! ")
        End If
    End If
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "ParseExecutionProfileSheet")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


''20230223: Modidfied to filter some data in Execution Profile Sheet.
Private Function IgnoreItem(sTemp As String) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If sTemp = vbNullString Or _
        sTemp = "NOP" Or _
        sTemp = "CALL" Or _
        sTemp = "PRINT" Or _
        sTemp = "FOR" Or _
        sTemp = "NEXT" Or _
        sTemp = "RETURN" Or _
        sTemp = "BINTABLE" Or _
        sTemp Like "*-*" Then
        
        IgnoreItem = True
    Else
    
        IgnoreItem = False
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "IgnoreItem") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''20230223: Modidfied to add new function to judge what instance need to plot current profile(instance in CurrentProfile_sheet and ExecutionProfileSheet the same).
' [20230809][All][Oliver] modify for current profile other sheet method
Public Function ProfileCollect_Start(TestInstan_Str As String)
    
    On Error GoTo errHandler
    
    Dim funcName As String:: funcName = "ProfileCollect_Start"
    
    Dim sTemp As String
    Dim sSubFlow As String
    Dim PinName_S As String
    Dim PlotTime_D As Double
    Dim s_SampleRate As String
    
    Dim s_PlotName As String
    Dim CurrentLine As Long
    s_PlotName = vbNullString
    CurrentLine = TheExec.Flow.CurrentLineNumber
    If gl_EnableVoltageProfile Or gl_EnableCurrentProfile Then

        If gl_isParExecutionProfileSheet = True Then
        
            sSubFlow = LCase(theexec.Flow.CurrentFlowSheetName)
            sTemp = LCase(sSubFlow & "_" & TestInstan_Str)
            If glb_CurrentProfile_Dict.Exists("all") Then
                s_PlotName = "all"
            ElseIf glb_CurrentProfile_Dict.Exists(sTemp) Then
                s_PlotName = sTemp
            ElseIf glb_CurrentProfile_Dict.Exists(sSubFlow) Then
                s_PlotName = sSubFlow
            Else
                Call Print_Error_Message(Error_Info, LIB_Common, "ProfileCollect_Start", "Error in wrong current profile sheet argument!")
            End If
        
            If s_PlotName <> "" Then
                If glb_ExecutionProfile_Dict.Exists(LCase(sTemp) & "_" & CurrentLine) Then

                    PlotTime_D = glb_ExecutionProfile_Dict.item(LCase(sTemp) & "_" & CurrentLine)
                    If glb_CurrentProfile_Dict.item(s_PlotName)(0) <> "" Then
                        PinName_S = glb_CurrentProfile_Dict.item(s_PlotName)(0)
                    Else
                        PinName_S = Core_Power_DCVS_pins + "," + Core_Power_DCVI_pins
                    End If
                    s_SampleRate = glb_CurrentProfile_Dict.item(s_PlotName)(1)
                    If s_SampleRate = "" Then s_SampleRate = 0
                    
                    If gl_EnableVoltageProfile Then
                        Start_Profile_AutoResolution PinName_S, "V", , PlotTime_D + 0.01, False, CLng(s_SampleRate)
                    ElseIf gl_EnableCurrentProfile Then
                        Start_Profile_AutoResolution PinName_S, "I", , PlotTime_D + 0.01, False, CLng(s_SampleRate)
                    End If
                Else
                    Call Print_Error_Message(Error_Info, LIB_Common, "ProfileCollect_Start", "Excution profile don't have this item!")
                End If
            End If
        End If
    End If
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "ProfileCollect_Start")
    If AbortTest Then Exit Function Else Resume Next
End Function


''20230223: Modidfied to add new function to judge what instance need to plot current profile(instance in CurrentProfile_sheet and ExecutionProfileSheet the same).
' [20230809][All][Oliver] modify for current profile other sheet method
Public Function ProfileCollect_End(TestInstan_Str As String)
    
    On Error GoTo errHandler
    
    Dim funcName As String:: funcName = "ProfileCollect_End"
    
    Dim sTemp As String
    Dim sSubFlow As String
    Dim PinName_PinL As New PinList
    Dim s_PinName As String
    Dim s_PlotName As String
    s_PlotName = vbNullString
    
    If gl_EnableVoltageProfile Or gl_EnableCurrentProfile Then
    
        If gl_isParExecutionProfileSheet = True Then
            sSubFlow = LCase(theexec.Flow.CurrentFlowSheetName)
            sTemp = LCase(sSubFlow & "_" & TestInstan_Str)
            
            If glb_CurrentProfile_Dict.Exists("all") Then
                s_PlotName = "all"
            ElseIf glb_CurrentProfile_Dict.Exists(sTemp) Then
                s_PlotName = sTemp
            ElseIf glb_CurrentProfile_Dict.Exists(sSubFlow) Then
                s_PlotName = sSubFlow
            Else
                Call Print_Error_Message(Error_Info, LIB_Common, "ProfileCollect_Start", "Error in wrong current profile sheet argument!")
            End If
            
            If s_PlotName <> "" Then
                If glb_ExecutionProfile_Dict.Exists(LCase(sTemp) & "_" & THEEXEC.Flow.CurrentLineNumber) And _
                    (gl_EnableCurrentProfile Or gl_EnableVoltageProfile) Then

                    If glb_CurrentProfile_Dict.item(s_PlotName)(0) <> "" Then
                        PinName_PinL.value = glb_CurrentProfile_Dict.item(s_PlotName)(0)
                    Else
                        PinName_PinL.value = Core_Power_DCVS_pins + "," + Core_Power_DCVI_pins  'PinName_S = "CorePower"
                    End If
                    Plot_Profile PinName_PinL, , True, False
                Else
                    Call Print_Error_Message(Error_Info, LIB_Common, "ProfileCollect_End", "Excution profile don't have this item!")
                End If
            End If
        End If
    End If
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "ProfileCollect_End")
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20230424][All][Tank] Parsing CurrentProfile_ sheet
' [20230809][All][Oliver] modify for get data by "GetSheetInfo" function
Public Function Get_CurrentProfile()
    
    Dim funcName As String:: funcName = "Get_CurrentProfile"
    
    Dim i As Long
    Dim j As Long
    Dim Count As Long: Count = 0
    Dim max_row As Long
    Dim max_col As Long
    
    Dim CurrentProfile_Sheet As String
    
    Dim sSubFlow As String: sSubFlow = "Subflow"
    Dim sInstance As String: sInstance = "Instance"
        
    Dim sheetNames() As String

    Dim vCurrentProfile() As Variant
    On Error GoTo errHandler
      
    Dim sTemp As String
    Dim n_StartValue_row As Integer
    Dim sSubflow_Temp As String
    Dim sInstance_Temp As String
    Dim n_Subflow_col As Integer
    Dim n_Instance_col As Integer
    Dim n_PowerPin_col As Integer
    Dim n_SampleSize_col As Integer
    
    Dim s_Subflow As String
    Dim s_Instance As String
    Dim s_PowerPin As String
    Dim s_SampleSize As String
    
    n_StartValue_row = 0
    n_Subflow_col = 0
    n_Instance_col = 0
    n_PowerPin_col = 0
    n_SampleSize_col = 0
    
    s_Subflow = vbNullString
    s_Instance = vbNullString
    s_PowerPin = vbNullString
    s_SampleSize = vbNullString
    CurrentProfile_Sheet = "CurrentProfile_"
    If Flag_GetCurrentProfile = False Then

        glb_CurrentProfile_Dict.RemoveAll
        
        If GetSheetInfo(CurrentProfile_Sheet, max_row, max_col, vCurrentProfile) Then
            If max_row > 0 And max_col > 0 Then
                '====process get header place====
                For i = 1 To max_row
                    If vCurrentProfile(i, 1) <> "" Then
                        For j = 1 To max_col
                            Select Case LCase(vCurrentProfile(i, j))
                                Case "subflow":
                                    n_Subflow_col = j
                                Case "instance":
                                    n_Instance_col = j
                                Case "powerpin":
                                    n_PowerPin_col = j
                                Case "samplerate":
                                    n_SampleSize_col = j
                            End Select
                        Next j
                        If n_Subflow_col <> 0 And n_Instance_col <> 0 Then
                            n_StartValue_row = i + 1
                            Exit For
                        End If
                    End If
                Next i
                
                If n_StartValue_row <> 0 Then
                '====process get header place====
                    For i = n_StartValue_row To max_row
                        s_Subflow = vbNullString
                        s_Instance = vbNullString
                        s_PowerPin = vbNullString
                        s_SampleSize = vbNullString
                        
                        If LCase(vCurrentProfile(i, n_Subflow_col)) = LCase("End") Then
                            Exit For
                        ElseIf vCurrentProfile(i, n_Subflow_col) <> "" Then
                            s_Subflow = vCurrentProfile(i, n_Subflow_col)
                            s_Instance = vCurrentProfile(i, n_Instance_col)
                            If n_PowerPin_col <> 0 Then s_PowerPin = vCurrentProfile(i, n_PowerPin_col)
                            If n_SampleSize_col <> 0 Then s_SampleSize = vCurrentProfile(i, n_SampleSize_col)
                            
                            If s_Instance <> "" Then
                                sTemp = LCase(s_Subflow & "_" & s_Instance)
                                If Not glb_CurrentProfile_Dict.Exists(sTemp) Then
                                    glb_CurrentProfile_Dict.Add sTemp, Array(s_PowerPin, s_SampleSize)
                                End If
                            Else
                                s_Subflow = LCase(s_Subflow)
                                If Not glb_CurrentProfile_Dict.Exists(s_Subflow) Then
                                    glb_CurrentProfile_Dict.Add s_Subflow, Array(s_PowerPin, s_SampleSize)
                                End If
                                
                                If LCase(vCurrentProfile(i, n_Subflow_col)) = LCase("All") Then Exit For
                            End If

                        End If
                    Next i
                End If
            End If
        End If
        Flag_GetCurrentProfile = True
    End If
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "Get_CurrentProfile")
    If AbortTest Then Exit Function Else Resume Next
End Function


''20230223: Modidfied to add new function let EVS can parallel ramp pin group.
' [20230906][T-Don][Tank] Modify EVS parallel plot current profile bug
' [20231228][T-All][Tank] Add EVS parallel IVCurve
Public Function Evs_Ramp_UPorDown_Parallel(Parallel_Pin_Voltage As String, Direction As String, d_PowerPin_PresentVol() As Double, S_WaitTime As Double, b_EVS_Printout As Boolean, Open_LatchUp_measure As Boolean, Optional Step_number As Integer = 10, Optional Rising_Delay_time As Double = 0.02, Optional TotalPwrLimit As Double = 20)

    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Evs_Ramp_UPorDown_Parallel_new"
    
    Dim d_StartVoltage As Double
''    d_StartVoltage = 0.75
    ''Parallel_Pin_Voltage = 1.1:VDD_SOC,VDD_GPU;1.0:VDDIO12
    Dim i As Long
    Dim j As Long
    Dim l_pincnt As Long
    Dim EVS_Index As Long
    
    Dim s_Pin_List As String
    Dim s_Pin_Ary() As String
    Dim s_GRP_Pin() As String
    Dim s_GRP_Pin_List() As String
    Dim s_GRP_Voltage_List() As String
    
    Dim d_GRP_Pin_StepSize() As Double
    
    Dim v_site As Variant
    Dim sb_Gatecheck() As New SiteBoolean
    Dim Gatecheck As Boolean
    
    Dim pld_Pin_Voltage As New PinListData
    
    If InStr(Parallel_Pin_Voltage, ";") <> 0 Then
        s_GRP_Pin = Split(Parallel_Pin_Voltage, ";")
    Else
        ReDim s_GRP_Pin(0)
        s_GRP_Pin(0) = Parallel_Pin_Voltage
    End If
    
    ReDim s_GRP_Pin_List(UBound(s_GRP_Pin))
    ReDim s_GRP_Voltage_List(UBound(s_GRP_Pin))
    
    ReDim d_GRP_Pin_StepSize(UBound(s_GRP_Pin))
    If UCase(Direction) = "UP" Then ReDim d_PowerPin_PresentVol(UBound(s_GRP_Pin))
    
    For i = 0 To UBound(s_GRP_Pin)
        s_GRP_Pin_List(i) = Split(s_GRP_Pin(i), ":")(1)
        s_GRP_Voltage_List(i) = Split(s_GRP_Pin(i), ":")(0)
        If UCase(Direction) = "UP" Then
            d_PowerPin_PresentVol(i) = FormatNumber(TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Voltage.value, 3)
        End If
''        d_StartVoltage = FormatNumber(thehdw.DCVS.Pins(s_GRP_Pin_List(i)).Voltage.value, 3)
        d_GRP_Pin_StepSize(i) = (CDbl(s_GRP_Voltage_List(i)) - d_PowerPin_PresentVol(i)) / Step_number
        
        If i = 0 Then
            s_Pin_List = s_GRP_Pin_List(i)
        Else
            s_Pin_List = s_Pin_List & "," & s_GRP_Pin_List(i)
        End If
    Next i
    
    theexec.DataManager.DecomposePinList s_Pin_List, s_Pin_Ary, l_pincnt
    
    Dim Power_Name As String
    Dim Latch_Up_name As String
    Dim m_InstanceName As String
    Dim Total_Power_Name As String
    
    Dim ForceV As Double
    Dim TotalPower As New SiteDouble
    Dim pld_PowerValue As New PinListData
    Dim pld_PowerPinVoltage As New PinListData
    Dim LatchUp_measure_Value As New PinListData
    
    Dim sa_Latch_Pin_Ary() As String
    Dim n_Latch_pin_cnt As Long
    Dim v_pin As Variant
    Dim s_pin As String
    Dim d_RampVoltage As Double
    Dim site As Variant 'Carter, 20240304
    m_InstanceName = theexec.DataManager.instancename
    
    ReDim sb_Gatecheck(l_pincnt - 1)
    If UCase(Direction) = "UP" Then
        For EVS_Index = 1 To Step_number
            For i = 0 To UBound(s_GRP_Pin)
                d_RampVoltage = d_PowerPin_PresentVol(i) + EVS_Index * d_GRP_Pin_StepSize(i)
                TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Voltage.value = d_RampVoltage
                If Open_LatchUp_measure Then
                    theexec.DataManager.DecomposePinList s_GRP_Pin_List(i), sa_Latch_Pin_Ary, n_Latch_pin_cnt
                    '' Print out measure current value if want to collect the Latch up data
                    If TheExec.enableWord("CurrentProfile") = True Or TheExec.Flow.enableWord("VoltageProfile") = True Or Profile_byflow = True Then
                        LatchUp_measure_Value.AddPin (s_GRP_Pin_List(i))
                        LatchUp_measure_Value.Pins(s_GRP_Pin_List(i)) = 0.001
                    Else
                        TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Meter.mode = tlDCVSMeterCurrent
                        LatchUp_measure_Value = TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                    End If
                    For Each v_pin In sa_Latch_Pin_Ary
                        s_pin = LCase(CStr(v_pin))
                        Latch_Up_name = m_InstanceName & "_" & "Latch_up_data_MeasI_" & Replace(CStr(s_GRP_Voltage_List(i)), ".", "p") & "V"
                        theexec.Flow.TestLimit resultVal:=LatchUp_measure_Value.Pins(s_pin), PinName:=s_pin, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Latch_Up_name, ForceVal:=d_RampVoltage, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                    Next v_pin
                End If
                TheHdw.Wait Rising_Delay_time
            Next i
        Next EVS_Index
        
        If b_EVS_Printout <> True Then
            For i = 0 To UBound(s_GRP_Pin)
                If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Or Profile_byflow Then
                    LatchUp_measure_Value.AddPin (s_GRP_Pin_List(i))
                    For Each site In theexec.sites
                        LatchUp_measure_Value.Pins(s_GRP_Pin_List(i)).value = 0.001
                    Next site
                Else
                    LatchUp_measure_Value = TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                End If
                ForceV = TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Voltage.value
                If pld_PowerValue.Pins.Count = 0 Then
                    pld_PowerValue = LatchUp_measure_Value.Math.Multiply(ForceV)
                Else
                    For j = 0 To LatchUp_measure_Value.Pins.Count - 1
                        pld_PowerValue.AddPin LatchUp_measure_Value.Pins(j).name
                        pld_PowerValue.Pins(LatchUp_measure_Value.Pins(j).name) = LatchUp_measure_Value.Pins(j).Multiply(ForceV)
                    Next j
                End If
    
                Latch_Up_name = m_InstanceName & "_" & "Latch_up_data_MeasI_" & Replace(CStr(Round(ForceV, 3)), ".", "p") & "V"
                theexec.Flow.TestLimit resultVal:=LatchUp_measure_Value, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Latch_Up_name, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
    
                Set LatchUp_measure_Value = Nothing
            Next i
        End If
        
        TheHdw.Wait S_WaitTime
        
        ''Printout Parallel EVS condition, Jayden 20221116
        If b_EVS_Printout <> True Then
            If glb_TesterType = "Jaguar" Then
                For i = 0 To UBound(s_Pin_Ary)
                    pld_PowerPinVoltage.AddPin s_Pin_Ary(i)
                    pld_PowerPinVoltage.Pins(s_Pin_Ary(i)) = TheHdw.DCVS.Pins(s_Pin_Ary(i)).Voltage.value
                Next i
            Else ''glb_TesterType = "UltraFLEXplus" Then
                #If IGXL_VER_1030 = True Then
                    pld_PowerPinVoltage = TheHdw.DCVS.Pins(s_Pin_List).Voltage.PinListData
                #Else
                    For i = 0 To UBound(s_Pin_Ary)
                        pld_PowerPinVoltage.AddPin s_Pin_Ary(i)
                        pld_PowerPinVoltage.Pins(s_Pin_Ary(i)) = TheHdw.DCVS.Pins(s_Pin_Ary(i)).Voltage.value
                    Next i
                #End If
            End If
            
            TotalPower = pld_PowerValue.Analyze.Sum
            
            theexec.Flow.TestLimit TotalPower, 0, TotalPwrLimit, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Total_Power_Name, PinName:="Total Power", customUnit:="W"     'BurstResult=1:Pass
            
            For Each site In theexec.sites
                If TotalPower(site) > TotalPwrLimit Then
                    theexec.sites(site).FlagState("F_EVS_POWER") = logicTrue
                Else
                    theexec.sites(site).FlagState("F_EVS_POWER") = logicFalse
                End If
            Next site
            
        End If

        theexec.Datalog.WriteComment ""
        theexec.Datalog.WriteComment "-------------------------EVS Power ramp start-------------------------"
        theexec.Flow.TestLimit S_WaitTime, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="EVS stress time", PinName:="EVS stress time", customUnit:="sec" 'BurstResult=1:Pass
         
         ''Printout Parallel EVS condition, Jayden 20221226
        If b_EVS_Printout <> True Then
            Power_Name = m_InstanceName & "_" & "ForceV"
            theexec.Flow.TestLimit pld_PowerPinVoltage, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Power_Name, customUnit:="V"     'BurstResult=1:Pass
            Power_Name = m_InstanceName & "_" & "Power"
            theexec.Flow.TestLimit pld_PowerValue, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Power_Name, customUnit:="W"     'BurstResult=1:Pass
        End If
         
    ElseIf UCase(Direction) = "DOWN" Then
        For EVS_Index = 1 To Step_number
            For i = 0 To UBound(s_GRP_Pin)
                d_RampVoltage = CDbl(s_GRP_Voltage_List(i)) - EVS_Index * d_GRP_Pin_StepSize(i)
                TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Voltage.value = d_RampVoltage
                If Open_LatchUp_measure Then
                    theexec.DataManager.DecomposePinList s_GRP_Pin_List(i), sa_Latch_Pin_Ary, n_Latch_pin_cnt
                    '' Print out measure current value if want to collect the Latch up data
                    If gl_EnableCurrentProfile = True Or gl_EnableVoltageProfile = True Or Profile_byflow = True Then
                        LatchUp_measure_Value.AddPin (s_GRP_Pin_List(i))
                        LatchUp_measure_Value.Pins(s_GRP_Pin_List(i)) = 0.001
                    Else
                        TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Meter.mode = tlDCVSMeterCurrent
                        LatchUp_measure_Value = TheHdw.DCVS.Pins(s_GRP_Pin_List(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                    End If
                    For Each v_pin In sa_Latch_Pin_Ary
                        s_pin = LCase(CStr(v_pin))
                        Latch_Up_name = m_InstanceName & "_" & "Latch_up_data_MeasI_" & Replace(CStr(s_GRP_Voltage_List(i)), ".", "p") & "V"
                        theexec.Flow.TestLimit resultVal:=LatchUp_measure_Value.Pins(s_pin), PinName:=s_pin, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Latch_Up_name, ForceVal:=d_RampVoltage, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                    Next v_pin
                End If
                TheHdw.Wait Rising_Delay_time
            Next i
        Next EVS_Index
    Else
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common.bas", "Evs_Ramp_UPorDown_Parallel", "Use not define status !!")
    End If
    
    For Each v_site In theexec.sites
        For i = 0 To UBound(s_Pin_Ary)
            sb_Gatecheck(i) = TheHdw.DCVS.Pins(s_Pin_Ary(i)).Gate
        Next i
    Next v_site
    
    
    If UCase(Direction) = "UP" Then
        For i = 0 To UBound(s_Pin_Ary)
            pld_Pin_Voltage.AddPin s_Pin_Ary(i)
            If sb_Gatecheck(i).Any(False) Then
                theexec.sites.Selected = sb_Gatecheck(i).LogicalXor(True)
                pld_Pin_Voltage.Pins(s_Pin_Ary(i)) = TheHdw.DCVS.Pins(s_Pin_Ary(i)).Voltage.value
                theexec.Flow.TestLimit pld_Pin_Voltage.Pins(s_Pin_Ary(i)), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V", PinName:=s_Pin_Ary(i), Tname:="Vlotage_PatternAlarm_After_Stress"
                theexec.Flow.TestLimit sb_Gatecheck(i), 1, 1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_Stress", PinName:=s_Pin_Ary(i)
                theexec.sites.Selected = True
            End If
        Next i
    
    ElseIf UCase(Direction) = "DOWN" Then
        For i = 0 To UBound(s_Pin_Ary)
            pld_Pin_Voltage.AddPin s_Pin_Ary(i)
            If sb_Gatecheck(i).Any(False) Then
                theexec.sites.Selected = sb_Gatecheck(i).LogicalXor(True)
                pld_Pin_Voltage.Pins(s_Pin_Ary(i)) = TheHdw.DCVS.Pins(s_Pin_Ary(i)).Voltage.value
                theexec.Flow.TestLimit sb_Gatecheck(i), -1, -1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_EVS_ramp_down", PinName:=s_Pin_Ary(i)
                theexec.sites.Selected = True
            End If
        Next i
        theexec.Datalog.WriteComment ""
        theexec.Datalog.WriteComment "--------------------------EVS Power ramp end--------------------------"
    Else
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common.bas", "Evs_Ramp_UPorDown_Parallel", "Use not define status !!")
    End If
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common.bas", "Evs_Ramp_UPorDown_Parallel")
    If AbortTest Then Exit Function Else Resume Next
End Function


''20230223: Modidfied to cheange subfunction "GetSheetInfo" place, and change sub -> boolean function.
Public Function GetSheetInfo(sheetName As String, MaxRow As Long, maxcol As Long, ParsingArr() As Variant) As Boolean

    Dim MySheet As New Worksheet
    Dim m_StartRow As Long:: m_StartRow = 1
    Dim sTemp As String
    
    On Error Resume Next

    Set MySheet = Sheets(sheetName)         ''Active sheet
    If err.number = 0 Then
    Else
        err.Clear
        GetSheetInfo = False
        Exit Function
    End If
    Application.ScreenUpdating = False
    
    MySheet.Activate

    MaxRow = MySheet.UsedRange.Rows.Count    ''Get the max used rows and columns of BDF.
    maxcol = MySheet.UsedRange.Columns.Count

    ParsingArr = MySheet.range(Cells(m_StartRow, 1), Cells(MaxRow, maxcol)).value
    
    Application.ScreenUpdating = True
    GetSheetInfo = True

End Function


' [20230407][T-All] Add to classify pin to DCVS and DCVI
' [20231228][T-All][Tank] Add classify pin to Digital
Public Function SortAllPinInstrumentType(sPinList As String, Optional sDCVS_Pin As String = vbNullString, Optional sDCVI_Pin As String = vbNullString, Optional sDigital_Pin As String = vbNullString)

Dim TempPinAry() As String
Dim nTempPinCnt As Long
Dim vPinName As Variant
Dim sPinName As String
Dim s_Msg As String
Dim s_TempPin As String

On Error GoTo errHandler

    theexec.DataManager.DecomposePinList sPinList, TempPinAry, nTempPinCnt
    
    For Each vPinName In TempPinAry
        sPinName = LCase(vPinName)
        If gl_GetInstrumentType_Dic.Exists(sPinName) Then
            If LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvs*" Then
                sDCVS_Pin = CombineStringList(sDCVS_Pin, sPinName)
            ElseIf LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvi*" Then
                sDCVI_Pin = CombineStringList(sDCVI_Pin, sPinName)
            ElseIf LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*i/o*" Then
                sDigital_Pin = CombineStringList(sDigital_Pin, sPinName)
            Else
                s_TempPin = CombineStringList(s_TempPin, sPinName)
            End If
        Else
            s_TempPin = CombineStringList(s_TempPin, sPinName)
        End If
        If s_TempPin <> "" Then
            s_Msg = "Pin : " & CStr(s_TempPin) & " are not define in DCVI and DCVS and Digital !!!"
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "SortAllPinInstrumentType", s_Msg)
        End If
    Next vPinName

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "SortAllPinInstrumentType")
    If AbortTest Then Exit Function Else Resume Next
End Function


'[20230407][T-All] Add to classify pin to different instrument type
Public Function GetAllPinInstrument(sPinList As String, sPin_HEXVS As String, sPin_VHDVS As String, sPin_VSM As String, sPin_VS800MA As String, sPin_VS5A As String, sPin_DC07 As String, sPin_DC30 As String, sPin_DC75 As String)

Dim TempPinAry() As String
Dim nTempPinCnt As Long
Dim sInputPin As String
Dim sTempPowerPinInstrument As String
Dim vPinName As Variant
Dim sPinName As String
Dim s_Msg As String

On Error GoTo errHandler
    theexec.DataManager.DecomposePinList sPinList, TempPinAry, nTempPinCnt

    If nTempPinCnt <> 0 Then
        For Each vPinName In TempPinAry
            sPinName = vPinName
            sInputPin = UCase(sPinName)
            If gl_GetInstrument_Dic.Exists(LCase(sInputPin)) Then
                sTempPowerPinInstrument = gl_GetInstrument_Dic(LCase(sInputPin))

                Select Case UCase(sTempPowerPinInstrument)
                    Case glbConstIns_HEXVS
                        sPin_HEXVS = CombineStringList(sPin_HEXVS, sInputPin)
                    Case glbConstIns_VHDVS
                        sPin_VHDVS = CombineStringList(sPin_VHDVS, sInputPin)
                    Case glbConstIns_VSM
                        sPin_VSM = CombineStringList(sPin_VSM, sInputPin)
                    Case glbConstIns_VS800MA
                        sPin_VS800MA = CombineStringList(sPin_VS800MA, sInputPin)
                    Case glbConstIns_VS5A
                        sPin_VS5A = CombineStringList(sPin_VS5A, sInputPin)
                    Case glbConstIns_DC07
                        sPin_DC07 = CombineStringList(sPin_DC07, sInputPin)
                    Case glbConstIns_DC30
                        sPin_DC30 = CombineStringList(sPin_DC30, sInputPin)
                    Case glbConstIns_DC75
                        sPin_DC75 = CombineStringList(sPin_DC75, sInputPin)
                    Case Else:
                        s_Msg = "Pin : " & sPinName & " isn't define in DCVI and DCVS !!!"
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common", "GetAllPinInstrument", s_Msg)
                End Select
            End If
        Next vPinName
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "GetAllPinInstrument")
    If AbortTest Then Exit Function Else Resume Next
End Function


'[20230407][T-All] Add to check array exist or not
Public Function Compare_Array_Exist(TempArray() As Variant) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If UBound(TempArray) <> 0 Then
        Compare_Array_Exist = True
    End If
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    err.Clear
    Compare_Array_Exist = False
End Function


'[20230407][RF] Add to do DCVI plot profile auto resolution
Public Function ProfileAutoResolution_DCVI(SlotType As String, measuretime As Double, ByRef sampleSize As Double, ByRef SampleRate As Double, Optional DownSampleRatio As Long = 1)

On Error GoTo errHandler

    Dim SampleSize_Ratio As Long: SampleSize_Ratio = 0
    Const dLowestSampleRate = 781.25
    
    '**************************UVI80***********************************
    Dim DC07MaxSampleSize As Double: DC07MaxSampleSize = 512000#
    Dim DC07MaxSampleRate As Double: DC07MaxSampleRate = 1000000#
    Dim DC07Maxtime As Double: DC07Maxtime = DC07MaxSampleSize / dLowestSampleRate
    
    'If UVI80_Mode = tlDCDiffMeterMode.tlDCDiffMeterModeHighAccuracy Then
    '    DC07MaxSampleSize = 1000#
    '    DC07MaxSampleRate = 100000#
    '    DC07Maxtime = DC07MaxSampleSize / 500
    'Else    'High Speed mode
    '    DC07MaxSampleSize = 4000000#
    '    DC07MaxSampleRate = 8330000#
    '    DC07Maxtime = DC07MaxSampleSize / 508.63
    'End If
    '******************************************************************
    
    '****************************DC-30***********************************
    Dim DC30MaxSampleSize As Double: DC30MaxSampleSize = 512#
    Dim DC30MaxSampleRate As Double: DC30MaxSampleRate = 100000#
    Dim DC30Maxtime As Double: DC30Maxtime = DC30MaxSampleSize / dLowestSampleRate
    '******************************************************************
    
    '****************************DC-75***********************************
    Dim DC75MaxSampleSize As Double: DC75MaxSampleSize = 512#
    Dim DC75MaxSampleRate As Double: DC75MaxSampleRate = 100000#
    Dim DC75Maxtime As Double: DC75Maxtime = DC75MaxSampleSize / dLowestSampleRate
    '******************************************************************
    
    Dim RealRate As Double
    Dim i As Integer
    Dim prediff As Double
    Dim posdiff As Double
    ''Time * SampleRate = SampleSize
    Select Case UCase(SlotType)
        Case glbConstIns_DC07
            If measuretime > DC07Maxtime Then
                sampleSize = DC07MaxSampleSize
                SampleRate = dLowestSampleRate
            Else
                If theexec.enableWord("DownSample_IProfile") = True Then
                    SampleSize_Ratio = DC07MaxSampleSize / DownSampleRatio
                    RealRate = SampleSize_Ratio / measuretime
                ElseIf SampleRate > 0 Then
                    RealRate = SampleRate
                Else
                    RealRate = DC07MaxSampleSize / measuretime
                End If
                'RealRate = DC07MaxSampleSize / measuretime
                If RealRate < dLowestSampleRate Then
                    SampleRate = dLowestSampleRate
                    sampleSize = DC07MaxSampleSize
                ElseIf RealRate > DC07MaxSampleRate Then
                    SampleRate = DC07MaxSampleRate
                    sampleSize = SampleRate * measuretime
                Else
                    SampleRate = RealRate
                    sampleSize = DC07MaxSampleSize
                End If
                
            End If
        Case glbConstIns_DC30
            If measuretime > DC30Maxtime Then
                sampleSize = DC30MaxSampleSize
                SampleRate = dLowestSampleRate
            Else
                If theexec.enableWord("DownSample_IProfile") = True Then
                    SampleSize_Ratio = DC30MaxSampleSize / DownSampleRatio
                    RealRate = SampleSize_Ratio / measuretime
                ElseIf SampleRate > 0 Then
                    RealRate = SampleRate
                Else
                    RealRate = DC30MaxSampleSize / measuretime
                End If
                'RealRate = DC30MaxSampleSize / measuretime
                If RealRate < dLowestSampleRate Then
                    SampleRate = dLowestSampleRate
                    sampleSize = DC07MaxSampleSize
                ElseIf RealRate > DC30MaxSampleRate Then
                    SampleRate = DC30MaxSampleRate
                    sampleSize = SampleRate * measuretime
                Else
                    SampleRate = RealRate
                    sampleSize = DC30MaxSampleSize
                End If
            End If
        Case glbConstIns_DC75
            If measuretime > DC75Maxtime Then
                sampleSize = DC75MaxSampleSize
                SampleRate = dLowestSampleRate
            Else
                If theexec.enableWord("DownSample_IProfile") = True Then
                    SampleSize_Ratio = DC75MaxSampleSize / DownSampleRatio
                    RealRate = SampleSize_Ratio / measuretime
                ElseIf SampleRate > 0 Then
                    RealRate = SampleRate
                Else
                    RealRate = DC75MaxSampleSize / measuretime
                End If
                'RealRate = DC75MaxSampleSize / measuretime
                If RealRate < dLowestSampleRate Then
                    SampleRate = dLowestSampleRate
                    sampleSize = DC07MaxSampleSize
                ElseIf RealRate > DC75MaxSampleRate Then
                    SampleRate = DC75MaxSampleRate
                    sampleSize = SampleRate * measuretime
                Else
                    SampleRate = RealRate
                    sampleSize = DC75MaxSampleSize
                End If
            End If
    End Select

    sampleSize = Ceiling(sampleSize)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "ProfileAutoResolution_DCVI")
    If AbortTest Then Exit Function Else Resume Next
    
End Function


' [20230420][All][Tank] change function place
Public Function GetDownSampleRatio(InstrType As String, ArySize As Integer) As Double

On Error GoTo errHandler
    Dim divide As Double
    Dim PinNumEachGroup As Integer: PinNumEachGroup = 10
   
    divide = Ceiling(ArySize / PinNumEachGroup)
    If divide >= 6 Then divide = 6
    
    If UCase(InstrType) = glbConstIns_VHDVS Then
        If divide >= 1 Then
            GetDownSampleRatio = 2
        Else
            GetDownSampleRatio = 1
        End If
    Else
        GetDownSampleRatio = 2 ^ divide
    End If
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "GetDownSampleRatio")
    If AbortTest Then Exit Function Else Resume Next
  
End Function


' [20230703][All][Tank] Exclude InitHi/Lo/HZ from digital pin no need to set 0v disconnect first
' [20231124][All][Tank] Check String not ""
Public Function ExcludePinFromPinList(OrgPinList As String, Exclude_Pin As String) As String
On Error GoTo errHandler

    Dim sa_OrgPinAry() As String
    Dim sa_PiniAry() As String
    
    Dim n_OrgPinCnt As Long
    Dim n_PinCnt As Long
    
    Dim v_DUTPin As Variant
    
    Dim dic_ExcludePin As New Dictionary
    
    If OrgPinList <> "" Then
        theexec.DataManager.DecomposePinList OrgPinList, sa_OrgPinAry(), n_OrgPinCnt
        
        If n_OrgPinCnt <> 0 Then
        
            If Exclude_Pin <> "" Then
                theexec.DataManager.DecomposePinList Exclude_Pin, sa_PiniAry(), n_PinCnt
                
                If n_PinCnt <> 0 Then
                    For Each v_DUTPin In sa_PiniAry
                        If dic_ExcludePin.Exists(LCase(CStr(v_DUTPin))) = False Then
                            dic_ExcludePin.Add LCase(CStr(v_DUTPin)), 0
                        End If
                    Next v_DUTPin
                    
                    For Each v_DUTPin In sa_OrgPinAry
                        If dic_ExcludePin.Exists(LCase(CStr(v_DUTPin))) = False Then
                            ExcludePinFromPinList = CombineStringList(ExcludePinFromPinList, CStr(v_DUTPin))
                        End If
                    Next v_DUTPin
                Else
                    ExcludePinFromPinList = OrgPinList
                End If
            Else
                ExcludePinFromPinList = OrgPinList
            End If
        Else
            ExcludePinFromPinList = vbNullString
        End If
    End If

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "ExcludePinFromPinList")
    If AbortTest Then Exit Function Else Resume Next
  
End Function


' [20230809][All][Oliver] add for current profile other sheet method
Private Function IgnoreFlow(sTemp As String) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If sTemp = vbNullString Or _
        sTemp Like "*_Init_Flag" Or _
        sTemp Like "*_Init_EnableWd" Then
        
        IgnoreFlow = True
    Else
    
        IgnoreFlow = False
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common", "IgnoreFlow") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230906][All][Carter] Modify fix process powerdown when Conti binout
' [20230907][T-Est][Oliver] Add alarmFail for initial
Public Function InitVariableOnStarted()
On Error GoTo errHandler
    glb_ApplyLevelTiming_FRC_Flag = False
    alarmFail = False
    glb_SFC_Scan_Check = False
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "InitVariableOnStarted")
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20230906][T-Spa][Jim] SBC Free Running Clock gating
Public Function FRC_Compare(PortName As String, Optional Freq As Double, Optional Threshold_Range As String = "1")
On Error GoTo errHandler

    Dim MeasClock As New PinListData
    Dim targetClock As Double
    Dim PinName As String
    Dim period As Double
    Dim DiffPinName() As String
    Dim i As Long
    Dim CounterValue As New PinListData
    Dim t_interval As New SiteDouble
    Dim CheckPins As String
    Dim SingleEndPin() As String
    Dim lng_CheckPinsNum As Long
    Dim Interval As Double
    Dim site As Variant 'Carter, 20240304
    Interval = 0.01
        
    If Threshold_Range = "" Then
        Threshold_Range = "1"
    End If
    
    If UCase(PortName) Like "*_DIFF*" Then                      'Diff
        Call theexec.DataManager.DecomposePinList(PortName, DiffPinName, lng_CheckPinsNum)

        For i = 0 To UBound(DiffPinName)
            
            If InStr(1, UCase(DiffPinName(i)), "_PA") <> 0 Then
                If UCase(DiffPinName(i)) Like "*XI*_PA" Then
                    
                    With TheHdw.Digital.Pins(DiffPinName(i)).FreqCtr
                        .EventSource = VOH
                        .EventSlope = Positive
                        .Enable = IntervalEnable
                        .Clear
                    End With
                    
                    TheHdw.Digital.Pins(DiffPinName(i)).FreqCtr.Interval = Interval
                    TheHdw.Digital.Pins(DiffPinName(i)).FreqCtr.start
                        
                    CounterValue = TheHdw.Digital.Pins(DiffPinName(i)).FreqCtr.Read
                    t_interval = TheHdw.Digital.Pins(DiffPinName(i)).FreqCtr.Interval
                    MeasClock = CounterValue.Math.divide(t_interval)
                                        
                    TheHdw.Digital.Pins(DiffPinName(i)).FreqCtr.Enable = Disable
                    
                    If Freq > 0 Then
                        targetClock = Freq
                    Else
                        period = TheHdw.Digital.Timing.period(PortName).value
                        targetClock = (1 / period)
                    End If
                    
                    For Each site In theexec.sites.Active
                        If MeasClock.Pins(DiffPinName(i)).Subtract(targetClock).divide(targetClock).Abs(site) > (Threshold_Range / 100) Then '20240703 michael added abs
                            Call Print_Error_Message(Error_Info, "LIB_Common", "FRC_Compare", "Wrong clock setup, Please use T-Autogen freerunning clock calculate tool")
                            theexec.sites.item(site).BinNumber = 15
                            theexec.sites.item(site).SortNumber = 9972
                            theexec.sites.item(site).result = tlResultFail
                        End If
                        theexec.Datalog.WriteComment "FRC_compare result site:" + CStr(site) + " targetclock : " + CStr(targetClock) + " Measureclock : " + CStr(MeasClock.Pins(DiffPinName(i)).value(site))
                    Next site
                End If
            End If
        Next i
    Else
        Call theexec.DataManager.DecomposePinList(PortName, SingleEndPin, lng_CheckPinsNum)
        For i = 0 To UBound(SingleEndPin)
            
            If InStr(1, UCase(SingleEndPin(i)), "_PA") <> 0 Then
                TheHdw.Digital.Pins(SingleEndPin(i)).Levels.value(chVoh) = (TheHdw.Digital.Pins(SingleEndPin(i)).Levels.value(chVih) * 0.9)
                With TheHdw.Digital.Pins(SingleEndPin(i)).FreqCtr
                    .EventSource = VOH
                    .EventSlope = Positive
                    .Enable = IntervalEnable
                    .Clear
                End With
                
                TheHdw.Digital.Pins(SingleEndPin(i)).FreqCtr.Interval = Interval
                TheHdw.Digital.Pins(SingleEndPin(i)).FreqCtr.start
                
                CounterValue = TheHdw.Digital.Pins(SingleEndPin(i)).FreqCtr.Read
                t_interval = TheHdw.Digital.Pins(SingleEndPin(i)).FreqCtr.Interval
                MeasClock = CounterValue.Math.divide(t_interval)
                                    
                TheHdw.Digital.Pins(SingleEndPin(i)).FreqCtr.Enable = Disable
        
                If Freq > 0 Then
                    targetClock = Freq
                Else
                    period = TheHdw.Digital.Timing.period(PortName).value
                    targetClock = (1 / period)
                End If
                
                For Each site In theexec.sites.Active
                    'If ((MeasClock.Pins(SingleEndPin(i)).value(site) - targetClock) / targetClock) > (Threshold_Range / 100) Then
                    If MeasClock.Pins(SingleEndPin(i)).Subtract(targetClock).divide(targetClock).Abs(site) > (Threshold_Range / 100) Then
                        Call Print_Error_Message(Error_Info, "LIB_Common", "FRC_Compare", "Wrong clock setup, Please use T-Autogen freerunning clock calculate tool")
                        theexec.sites.item(site).BinNumber = 15
                        theexec.sites.item(site).SortNumber = 9972
                        theexec.sites.item(site).result = tlResultFail
                    End If
                    theexec.Datalog.WriteComment "FRC_compare result site:" + CStr(site) + " targetclock : " + CStr(targetClock) + " Measureclock : " + CStr(MeasClock.Pins(SingleEndPin(i)).value(site))
    
                Next site
            End If
        Next i

        
    End If

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "FRC_Compare")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function PPMUInit(s_pin As String, d_InputValue As Double, s_ForceType As String)
On Error GoTo errHandler

    If s_pin <> "" Then
        TheHdw.Digital.Pins(s_pin).Disconnect
        With TheHdw.PPMU.Pins(s_pin)
            If LCase(s_ForceType) = "v" Then
                .ForceV d_InputValue
            ElseIf LCase(s_ForceType) = "i" Then
                .ForceI d_InputValue
            Else
            End If
            .Connect
            .Gate = tlOn
        End With
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common.bas", "PPMUInit")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][HY] Just test WalkingZ fail pins
Public Function isContiTestFailPinList(s_OrgPinList As String, dic_ComparePinList As Dictionary, s_NewPinLsit As String) As Boolean
    Dim sa_Pins() As String
    Dim n_Pin_cnt As Long
    Dim n As Long
On Error GoTo errHandler
    isContiTestFailPinList = False
    theexec.DataManager.DecomposePinList s_OrgPinList, sa_Pins(), n_Pin_cnt
    For n = 0 To UBound(sa_Pins)
        If dic_ComparePinList.Exists(LCase(sa_Pins(n))) Then
            s_NewPinLsit = CombineStringList(s_NewPinLsit, sa_Pins(n))
            isContiTestFailPinList = True
        End If
    Next n
 Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "isContiTestFailPinList")
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20231106][T-Tah][Oliver] add multiple SRAM pin compares with one logic pin method
Public Function SelSRAM_Index_Select(SelsramMapping() As SELSRAM_Bit_Table, Optional s_block As String, Optional s_pattern As String) As Long
    On Error GoTo errHandler
    
    Dim k As Long
    Dim i As Long
    Dim PattIdx As Long
    Dim DuplicateBlockExist As Boolean: DuplicateBlockExist = False
    PattIdx = -1
    
    For i = 0 To UBound(StrAry_DuplicateBlock)
        If UCase(s_block) Like StrAry_DuplicateBlock(i) Then DuplicateBlockExist = True
    Next i
    If DuplicateBlockExist = True Then
        For k = 0 To UBound(SelsramMapping)
            If (UCase(s_block) Like UCase(SelsramMapping(k).blockName)) And (SelsramMapping(k).DuplicateBlock = True) Then
                If UCase(SelsramMapping(k).blockName) <> "*" And (s_block <> "") And (UCase(s_block) Like UCase(SelsramMapping(k).blockName)) _
                    And UCase(SelsramMapping(k).Pattern) <> "*" And (s_pattern <> "") And (UCase(s_pattern) Like UCase(SelsramMapping(k).Pattern)) Then
                    PattIdx = k
                    Exit For
                End If
            End If
        Next k
    Else
        For k = 0 To UBound(SelsramMapping)
            If UCase(SelsramMapping(k).blockName) <> "*" And (s_block <> "") And (UCase(s_block) Like UCase(SelsramMapping(k).blockName)) Then
                PattIdx = k
                Exit For
            ElseIf UCase(SelsramMapping(k).Pattern) <> "*" And (s_pattern <> "") And (UCase(s_pattern) Like UCase(SelsramMapping(k).Pattern)) Then
                PattIdx = k
                Exit For
'            Else
'                theexec.Datalog.WriteComment "Warning: Could not find the mapping on the SELSRM_Mapping_Table."
            End If
        Next k
    End If
    SelSRAM_Index_Select = PattIdx
    
    Exit Function
                                                                                                                                                                                                                                                               
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Index_Select")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Decide_Switching_Bit(digSrc_EQ As String, DSPWaveSwitch As DSPWave, DC_Level As PinListData, BlockType As String, SELSRM_Rails As String, _
                                    Optional shmoo_pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary, _
                                    Optional DSSCPatName As String, Optional set_init As Boolean, Optional seq As Long, Optional Power_Run_Scenario As String, _
                                    Optional SEL_String As SiteVariant, Optional printout As Boolean = False) As String
  
    On Error GoTo errHandler
    
    Dim v_site As Variant
    Dim p_cnt As Long
    Dim p_ary() As String
    Dim s_alpha As String
    Dim s_logicPin As String
    Dim s_SramPin As String
    Dim DSSC_Switching_Voltage As New PinListData

    Dim DSSCSelSrmOpposite As Long
    Dim BlockTypeNum As Long
    Dim PattIdx As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim s_Statement As String
    Dim ReturnString() As String
    Dim sv_LogicValue As New SiteVariant
    Dim sv_SramValue As New SiteVariant
    
    Dim sb_Sdomain As New SiteBoolean
    Dim sl_Sdomain As New SiteLong
    Dim sl_ReturnString() As New SiteLong
    
    Dim pld_LogicValue As New PinListData
    Dim pld_SramValue As New PinListData
    'Dim tmpString As New SiteVariant
    
    BlockTypeNum = -1
    PattIdx = -1
    ReDim ReturnString(Len(digSrc_EQ) - 1)
    ReDim sl_ReturnString(Len(digSrc_EQ) - 1)
    Decide_DSSC_Switching_Voltage DSSC_Switching_Voltage, DC_Level, shmoo_pin, ShmooPinsVoltage, ForcePin, SetForceVoltage, set_init, seq, Power_Run_Scenario
    
    Dim l_Selsram_index As Long
    l_Selsram_index = SelSRAM_Index_Select(SelsramMapping, BlockType, DSSCPatName)
   
    If l_Selsram_index <> -1 Then
        Call SelSRAM_DigSrc_Bit(l_Selsram_index, digSrc_EQ, DSSC_Switching_Voltage, DSPWaveSwitch, ReturnString, sl_ReturnString, , SEL_String, printout)
        If theexec.DevChar.Setups.IsRunning = False Then
            Decide_Switching_Bit = Join(ReturnString, "")
            SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit, BlockType, DSSCPatName)
        Else
            Decide_Switching_Bit = digSrc_EQ
            SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit, BlockType, DSSCPatName)
        End If
        
    Else
        s_Statement = "We could not find from the SelSRAM Mapping Table properly."
        ''Standardize error message
    End If
 
  Exit Function
  
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "Decide_Switching_Bit")
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20231106][T-Tah][Oliver] add multiple SRAM pin compares with one logic pin method
Public Function SelSRAM_DigSrc_Bit(l_Selsram_index As Long, s_DSSC_EQ_Assignment As String, pld_DSSC_Switching_Voltage As PinListData, _
            dsp_DSSC_Switch As DSPWave, s_ReturnString() As String, sl_ReturnString() As SiteLong, _
            Optional b_HIP_SelSram As Boolean, Optional s_CombineString As SiteVariant, Optional printing As Boolean) As SiteLong
    
    On Error GoTo errHandler

    Dim v_site As Variant
    Dim p_cnt As Long
    Dim p_ary() As String
    Dim s_alpha As String
    Dim s_logicPin As String
    Dim s_SramPin As String

    Dim DSSCSelSrmOpposite As Long
    Dim BlockTypeNum As Long
    Dim PattIdx As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim PowerDomain_Idx As Long
    
    Dim sb_Sdomain As New SiteBoolean
    Dim sl_Sdomain As New SiteLong

    Dim pld_LogicValue As New PinListData
    Dim pld_SramValue As New PinListData
    
    'Dim s_CombineString As New SiteVariant

    If b_HIP_SelSram Then
        For i = 0 To UBound(SelsramMapping(l_Selsram_index).digsrc_assignment)
            If UCase(SelsramMapping(l_Selsram_index).digsrc_assignment(i)) Like UCase(s_DSSC_EQ_Assignment) Then
                s_alpha = SelsramMapping(l_Selsram_index).alpha(i)
                s_logicPin = SelsramMapping(l_Selsram_index).logic_Pin(i) ''PinA
                s_SramPin = SelsramMapping(l_Selsram_index).sram_Pin(i) ''PinB;PinC
                DSSCSelSrmOpposite = SelsramMapping(l_Selsram_index).SelSrm1(i)
                PowerDomain_Idx = VddBinStr2Enum(s_logicPin)
                If UCase(s_logicPin) Like "PRESERVED" Then
                    dsp_DSSC_Switch.Element(i) = DSSCSelSrmOpposite
                    s_ReturnString(i) = DSSCSelSrmOpposite
                Else
                    Call SelSRAM_Voltage_To_PLD(s_logicPin, pld_DSSC_Switching_Voltage, pld_LogicValue)
                    Call SelSRAM_Voltage_To_PLD(s_SramPin, pld_DSSC_Switching_Voltage, pld_SramValue)
                    
''                    theexec.DataManager.DecomposePinList s_logicPin, p_ary, p_cnt
''                    For j = 0 To p_cnt - 1
''                        If gl_GetInstrumentType_Dic.Exists(LCase(p_ary(j))) Then
''                            pld_LogicValue.AddPin p_ary(j)
''                            pld_LogicValue.Pins(p_ary(j)) = pld_DSSC_Switching_Voltage.Pins(p_ary(j))
''                        End If
''                    Next j
''
''                    theexec.DataManager.DecomposePinList s_SramPin, p_ary, p_cnt
''                    For j = 0 To p_cnt - 1
''                        If gl_GetInstrumentType_Dic.Exists(LCase(p_ary(j))) Then
''                            pld_SramValue.AddPin p_ary(j)
''                            pld_SramValue.Pins(p_ary(j)) = pld_DSSC_Switching_Voltage.Pins(p_ary(j))
''                        End If
''                    Next j
                    
                    Call SelSRAM_Voltage_Comparison(pld_LogicValue, pld_SramValue, DSSCSelSrmOpposite, sl_Sdomain)
                    For Each v_site In theexec.sites
                        Set sl_ReturnString(PowerDomain_Idx) = sl_Sdomain
                    Next v_site
                End If
                Set pld_LogicValue = Nothing
                Set pld_SramValue = Nothing
                Exit For
            End If
        Next i
        
    Else
        For i = 0 To Len(s_DSSC_EQ_Assignment) - 1
            If UCase(CStr(mid(s_DSSC_EQ_Assignment, i + 1, 1))) Like "S" Then
                s_alpha = SelsramMapping(l_Selsram_index).alpha(i)
                s_logicPin = SelsramMapping(l_Selsram_index).logic_Pin(i) ''PinA
                s_SramPin = SelsramMapping(l_Selsram_index).sram_Pin(i) ''PinB;PinC
                DSSCSelSrmOpposite = SelsramMapping(l_Selsram_index).SelSrm1(i)
                If UCase(s_logicPin) Like "PRESERVED" Then
                    dsp_DSSC_Switch.Element(i) = DSSCSelSrmOpposite
                    s_ReturnString(i) = DSSCSelSrmOpposite
                                        sl_Sdomain = DSSCSelSrmOpposite
                    Set sl_ReturnString(i) = sl_Sdomain
                Else
                    Call SelSRAM_Voltage_To_PLD(s_logicPin, pld_DSSC_Switching_Voltage, pld_LogicValue)
                    Call SelSRAM_Voltage_To_PLD(s_SramPin, pld_DSSC_Switching_Voltage, pld_SramValue)
                    
''                    theexec.DataManager.DecomposePinList s_logicPin, p_ary, p_cnt
''                    For j = 0 To p_cnt - 1
''                        If gl_GetInstrumentType_Dic.Exists(LCase(p_ary(j))) Then
''                            pld_LogicValue.AddPin p_ary(j)
''                            pld_LogicValue.Pins(p_ary(j)) = pld_DSSC_Switching_Voltage.Pins(p_ary(j))
''                        End If
''                    Next j
''
''                    theexec.DataManager.DecomposePinList s_SramPin, p_ary, p_cnt
''                    For j = 0 To p_cnt - 1
''''                        If j >= 1 Then Stop
''                        If gl_GetInstrumentType_Dic.Exists(LCase(p_ary(j))) Then
''                            pld_SramValue.AddPin p_ary(j)
''                            pld_SramValue.Pins(p_ary(j)) = pld_DSSC_Switching_Voltage.Pins(p_ary(j))
''                            If j >= 1 Then pld_SramValue.Pins(p_ary(j)) = 0.6
''                        End If
''                    Next j
                    
                    Call SelSRAM_Voltage_Comparison(pld_LogicValue, pld_SramValue, DSSCSelSrmOpposite, sl_Sdomain)
                    For Each v_site In theexec.sites
                        dsp_DSSC_Switch.ElementLite(i) = sl_Sdomain
                        Set sl_ReturnString(i) = sl_Sdomain
                    Next v_site
                    For Each v_site In theexec.sites
                        s_ReturnString(i) = CStr(sl_Sdomain(v_site))
                        Exit For
                    Next v_site
                End If
            Else

				For Each v_site In TheExec.sites
                    dsp_DSSC_Switch.Element(i) = CDbl(mid(s_DSSC_EQ_Assignment, i + 1, 1))
                Next v_site
                s_ReturnString(i) = CDbl(mid(s_DSSC_EQ_Assignment, i + 1, 1))
                sl_Sdomain(i) = CLng(mid(s_DSSC_EQ_Assignment, i + 1, 1))
                Set sl_ReturnString(i) = sl_Sdomain

            End If
            Set pld_LogicValue = Nothing
            Set pld_SramValue = Nothing
            If printing = True Then
                For Each v_site In theexec.sites
                        s_CombineString(v_site) = s_CombineString(v_site) & CStr(sl_ReturnString(i)(v_site))
                Next v_site
            End If
        Next i
    End If
    'SelSRAM_DigSrc_Bit = s_CombineString

    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_DigSrc_Bit")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231106][T-Tah][Oliver] add multiple SRAM pin compares with one logic pin method
' [20231228][T-All][Tank] Fix use global variable "site" change to local variable "v_site"
Public Function SelSRAM_Voltage_Comparison(ByVal pld_LogicValue As PinListData, ByVal pld_SramValue As PinListData, ByVal DSSCSelSrmOpposite As Long, sl_Sdomain As SiteLong)
    On Error GoTo errHandler
    
    Dim i As Long
    Dim l_sram_cnt As Long
    Dim l_logic_cnt As Long
    Dim sd_sram_Ary() As New SiteDouble
    Dim sd_logic_Ary() As New SiteDouble
    
    Dim pin As Variant
    Dim v_site As Variant
    Dim b_ErrorLog As Boolean
    
    Dim sl_Sdomain_Temp As New SiteLong
    Dim sb_Sdomain As New SiteBoolean
    Dim sb_Sdomain_Temp As New SiteBoolean
    
    b_ErrorLog = False
    l_sram_cnt = pld_SramValue.Pins.Count
    l_logic_cnt = pld_LogicValue.Pins.Count
    ReDim sd_sram_Ary(l_sram_cnt - 1)
    ReDim sd_logic_Ary(l_logic_cnt - 1)
    
    Dim s_Statement As String
    For i = 0 To l_logic_cnt - 1
        For Each v_site In theexec.sites
            sd_logic_Ary(i) = pld_LogicValue.Pins(i).value
        Next
    Next i
    For i = 0 To l_sram_cnt - 1
        For Each v_site In theexec.sites
            sd_sram_Ary(i) = pld_SramValue.Pins(i).value
        Next
    Next i
    
    If DSSCSelSrmOpposite = 0 Then
        If (l_logic_cnt = 1) And (l_sram_cnt = 1) Then
            For i = 0 To l_logic_cnt - 1
                sb_Sdomain = sd_logic_Ary(i).compare(GreaterThan, sd_sram_Ary(i))
                sl_Sdomain = sb_Sdomain.If(1#, 0#)
            Next i
        ElseIf (l_logic_cnt = 1) And (l_sram_cnt > 1) Then
            ''Example: Log_PinA=[0.9,0.9]
            ''         Sram_PinB = [0.85, 0.85]; Sram_PinC = [0.95, 0.95][0.85, 0.85]; Sram_PinD = [0.85, 0.85]
            For i = 0 To l_sram_cnt - 1
                sb_Sdomain_Temp = sd_logic_Ary(0).compare(GreaterThan, sd_sram_Ary(i))
                ''1st, Log_PinA compare with Sram_PinB ---> sb_Sdomain_Temp = [True,True]
                ''2nd, Log_PinA compare with Sram_PinC ---> sb_Sdomain_Temp = [False,False]
                ''3rd, Log_PinA compare with Sram_PinD ---> sb_Sdomain_Temp = [True,True]
                If i = 0 Then
                    sl_Sdomain_Temp = sb_Sdomain_Temp.If(1#, 0#)
                    ''1st, Log_PinA compare with Sram_PinB ---> sl_Sdomain_Temp = [1,1]
                Else
                    sb_Sdomain = sl_Sdomain_Temp.compare(EqualTo, sb_Sdomain_Temp.If(1#, 0#))
                    ''2nd, Log_PinA compare with Sram_PinC ---> sl_Sdomain_Temp = [1,1] and sb_Sdomain_Temp = [0,0] --> sb_Sdomain = [False,False], b_ErrorLog = True
                    ''3rd, Log_PinA compare with Sram_PinD ---> sl_Sdomain_Temp = [1,1] and sb_Sdomain_Temp = [1,1] --> sb_Sdomain = [True,True]
                    If sb_Sdomain.Any(False) Then
''                        TheExec.Datalog.WriteComment "--------> Voltage Comparison has different result between Mutli SelSRAM Pins. Send out 0 as default."
                        ''Standardize error message
                        s_Statement = "Voltage Comparison has different result between Mutli SelSRAM Pins. Send out 0 as default."
                        Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Voltage_Comparison", s_Statement)
                        b_ErrorLog = True
                    End If
                End If
            Next i
            If b_ErrorLog Then
                sl_Sdomain = 0#
            Else
                sl_Sdomain = sl_Sdomain_Temp
            End If
        ElseIf ((l_logic_cnt > 1) And (l_sram_cnt = 1)) Then
''            TheExec.Datalog.WriteComment "--------> Currently, we do not have this scenario with multi logic pins and single sram pin."
            ''Standardize error message
            s_Statement = "Currently, we do not have this scenario with multi logic pins and single sram pin."
            Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Voltage_Comparison", s_Statement)
            sl_Sdomain = 0#
        Else
''            TheExec.Datalog.WriteComment "--------> Currently, we do not have this scenario with multi logic pins and multi sram pins."
            ''Standardize error message
            s_Statement = "Currently, we do not have this scenario with multi logic pins and single sram pin."
            Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Voltage_Comparison", s_Statement)
            sl_Sdomain = 0#
        End If
''        sb_Sdomain = sv_LogicValue.Compare(GreaterThan, sv_SramValue)
''        sl_Sdomain = sb_Sdomain.If(1#, 0#)

    ElseIf DSSCSelSrmOpposite = 1 Then
        If (l_logic_cnt = 1) And (l_sram_cnt = 1) Then
            For i = 0 To l_logic_cnt - 1
                sb_Sdomain = sd_logic_Ary(i).compare(GreaterThan, sd_sram_Ary(i))
                sl_Sdomain = sb_Sdomain.If(0#, 1#)
            Next i
        ElseIf (l_logic_cnt = 1) And (l_sram_cnt > 1) Then
            ''Example: Log_PinA=[0.9,0.9]
            ''         Sram_PinB = [0.85, 0.85]; Sram_PinC = [0.95, 0.95][0.85, 0.85]; Sram_PinD = [0.85, 0.85]
            For i = 0 To l_sram_cnt - 1
                sb_Sdomain_Temp = sd_logic_Ary(0).compare(GreaterThan, sd_sram_Ary(i))
                ''1st, Log_PinA compare with Sram_PinB ---> sb_Sdomain_Temp = [True,True]
                ''2nd, Log_PinA compare with Sram_PinC ---> sb_Sdomain_Temp = [False,False]
                ''3rd, Log_PinA compare with Sram_PinD ---> sb_Sdomain_Temp = [True,True]
                If i = 0 Then
                    sl_Sdomain_Temp = sb_Sdomain_Temp.If(0#, 1#)
                    ''1st, Log_PinA compare with Sram_PinB ---> sl_Sdomain_Temp = [1,1]
                Else
                    sb_Sdomain = sl_Sdomain_Temp.compare(EqualTo, sb_Sdomain_Temp.If(0#, 1#))
                    ''2nd, Log_PinA compare with Sram_PinC ---> sl_Sdomain_Temp = [1,1] and sb_Sdomain_Temp = [0,0] --> sb_Sdomain = [False,False], b_ErrorLog = True
                    ''3rd, Log_PinA compare with Sram_PinD ---> sl_Sdomain_Temp = [1,1] and sb_Sdomain_Temp = [1,1] --> sb_Sdomain = [True,True]
                    If sb_Sdomain.Any(False) Then
''                        TheExec.Datalog.WriteComment "--------> Voltage Comparison has different result between Mutli SelSRAM Pins. Send out 0 as default."
                        ''Standardize error message
                        s_Statement = "Voltage Comparison has different result between Mutli SelSRAM Pins. Send out 0 as default."
                        Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Voltage_Comparison", s_Statement)
                        b_ErrorLog = True
                    End If
                End If
            Next i
            If b_ErrorLog Then
                sl_Sdomain = 0#
            Else
                sl_Sdomain = sl_Sdomain_Temp
            End If
        ElseIf ((l_logic_cnt > 1) And (l_sram_cnt = 1)) Then
''            TheExec.Datalog.WriteComment "--------> Currently, we do not have this scenario with multi logic pins and single sram pin."
            ''Standardize error message
            s_Statement = "Currently, we do not have this scenario with multi logic pins and single sram pin."
            Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Voltage_Comparison", s_Statement)
            sl_Sdomain = 0#
        Else
''            TheExec.Datalog.WriteComment "--------> Currently, we do not have this scenario with multi logic pins and multi sram pins."
            ''Standardize error message
            s_Statement = "Currently, we do not have this scenario with multi logic pins and single sram pin."
            Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Voltage_Comparison", s_Statement)
            sl_Sdomain = 0#
        End If
''        sb_Sdomain = sv_LogicValue.Compare(GreaterThan, sv_SramValue)
''        sl_Sdomain = sb_Sdomain.If(0#, 1#)

    End If
    
    ''Datalog the error message if voltage is different between pins.
    If b_ErrorLog Then
        theexec.Datalog.WriteComment "--------> Logic Pin Voltage"
        For Each v_site In theexec.sites
            For Each pin In pld_LogicValue.Pins
                theexec.Datalog.WriteComment "--------> Site(" & v_site & "), Logic Pin(" & CStr(pin) & "): " & pld_LogicValue.Pins(pin).value
            Next pin
        Next v_site
        theexec.Datalog.WriteComment "--------> Sram Pin Voltage"
        For Each v_site In theexec.sites
            For Each pin In pld_SramValue.Pins
                theexec.Datalog.WriteComment "--------> Site(" & v_site & "), Sram Pin(" & CStr(pin) & "): " & pld_SramValue.Pins(pin).value
            Next pin
            theexec.sites.item(v_site).SortNumber = 9972
            theexec.sites.item(v_site).BinNumber = 15
            theexec.sites.item(v_site).result = tlResultFail
        Next v_site
    End If
    
    Exit Function
                                                                                                                                                                                                                                                               
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Voltage_Comparison")
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20231106][T-Tah][Oliver] add multiple SRAM pin compares with one logic pin method
Public Function SelSRAM_Voltage_To_PLD(ByVal s_logic_sram_Pin As String, ByVal pld_DSSC_Switching_Voltage As PinListData, pld_logic_sram_Voltage As PinListData)
    On Error GoTo errHandler
    Dim j As Long
    Dim p_cnt As Long
    Dim p_ary() As String
    
    theexec.DataManager.DecomposePinList s_logic_sram_Pin, p_ary, p_cnt
    For j = 0 To p_cnt - 1
        If gl_GetInstrumentType_Dic.Exists(LCase(p_ary(j))) Then
            pld_logic_sram_Voltage.AddPin p_ary(j)
            pld_logic_sram_Voltage.Pins(p_ary(j)) = pld_DSSC_Switching_Voltage.Pins(p_ary(j))
        End If
    Next j
    
    Exit Function
                                                                                                                                                                                                                                                               
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "SelSRAM_Voltage_To_PLD")
    If AbortTest Then Exit Function Else Resume Next
End Function


' [20231228][T-Son][Tank] Add Finger_Print need setup HRAM When pre pattern (if PLUS and set long function "short")
Public Function HRAMSetupInterpose(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    TheHdw.Digital.hram.size = gl_HRAMmaxDepth
    TheHdw.Digital.hram.CaptureType = captFail
    
    TheHdw.Digital.hram.SetTrigger trigFail, False, 0, True
    TheHdw.Digital.Patgen.ClearFail
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_DC", "HRAMSetupInterpose") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20240110][All][Tank] add store min Ifold
Public Function GetMinIFoldLimit(s_PowerPin As String, d_MinCurrentRange As Double) As Double
On Error GoTo errHandler
    Dim d_OrgCurrentRange As Double
    If gl_GetInstrumentType_Dic.Exists(LCase(s_PowerPin)) Then
        If UCase(gl_GetInstrumentType_Dic(LCase(s_PowerPin))) Like "*DCVS*" Then
            ' Store current range
            d_OrgCurrentRange = TheHdw.DCVS.Pins(s_PowerPin).CurrentRange.value
            ' Set min current range
            TheHdw.DCVS.Pins(s_PowerPin).SetCurrentRanges d_MinCurrentRange, d_MinCurrentRange
            ' Get min IFold limit value of min current range
            GetMinIFoldLimit = TheHdw.DCVS.Pins(s_PowerPin).CurrentLimit.Source.FoldLimit.level.min
            ' Restore current range
            TheHdw.DCVS.Pins(s_PowerPin).SetCurrentRanges d_OrgCurrentRange, d_OrgCurrentRange
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "GetMinIFoldLimit")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20240223][T-All][Clyde] Function Create to combine pin array to pin list string
Public Function PinAryToString(pin() As String) As String
On Error GoTo errHandler
    Dim i As Integer
    
    PinAryToString = ""
    For i = 0 To UBound(pin)
        PinAryToString = CombineStringList(PinAryToString, pin(i))
    Next i
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "PinAryToString")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20240216][All][Clyde, Steven] add binout function with VBA
Public Function BinoutSite(Optional ByVal site As Long = -1, Optional ByVal sort_bin As Long = -1, Optional ByVal hard_bin As Long = -1, Optional ByVal Fail_Flag As String = vbNullString, _
                            Optional ByVal binout_message As String = vbNullString)
On Error GoTo errHandler
    
    Dim v_site As Variant
    
    If site = -1 Then
        For Each v_site In theexec.sites
            Call BinoutSite(v_site, sort_bin, hard_bin, Fail_Flag, binout_message)
        Next v_site
    
    Else
        If sort_bin <> -1 Then
            If sort_bin <= 0 Then
                Call Print_Error_Message(Error_Info, "LIB_Common", "BinoutSite", "Sort bin number error, is " & sort_bin & ", change Sort bin number to " & glb_SortBin_Bin0)
                sort_bin = glb_SortBin_Bin0
            End If
            theexec.sites.item(site).SortNumber = sort_bin
        End If
        
        If hard_bin <> -1 Then
            If hard_bin <= 0 Then
                Call Print_Error_Message(Error_Info, "LIB_Common", "BinoutSite", "Hard bin number error, is " & hard_bin & ", change Hard bin number to " & glb_HardBin_Bin0)
                hard_bin = glb_HardBin_Bin0
            End If
            theexec.sites.item(site).BinNumber = hard_bin
        End If
        
        If Fail_Flag <> "" Then
            theexec.sites.item(site).FlagState(Fail_Flag) = logicTrue
        End If
        
        If binout_message <> "" Then
            theexec.Datalog.WriteComment "site:" & site & ", " & binout_message
        End If
        
        theexec.sites.item(site).result = tlResultFail
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "BinoutSite")
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20240311][T-All][Clyde] Add for Power Up and Power Down used, setup IO sequence number
Public Function PowerUpDownIOSeq(ioPins As String)
On Error GoTo errHandler
    Dim ioCnt As Long
    Dim ioPin() As String
    Dim pwrGloName As String
    Dim PinIdx As Integer
    Dim ioInitIdx As Integer
    Dim seqnum() As Integer
    Dim seqPin() As String
    Dim ioInit(2) As String
    Dim seqCnt As Integer
    Dim havePwrDown As Boolean
    Dim i As Integer
    
    functionName = "PowerUpDownIOSeq"
    ioInit(0) = IOInitHi
    ioInit(1) = IOInitLo
    ioInit(2) = IOInitHZ
    
    If ioPins <> "" Then
        theexec.DataManager.DecomposePinList ioPins, ioPin(), ioCnt
        ' power up
        For ioInitIdx = 0 To UBound(ioInit)
            seqCnt = 0
            For PinIdx = 0 To ioCnt - 1
                If theexec.DataManager.ChannelType(ioPin(PinIdx)) <> "N/C" Then
                    pwrGloName = ioPin(PinIdx) + PowerUpSeqName + ioInit(ioInitIdx)
                    If nonPowerSeqNum.Exists(UCase(pwrGloName)) Then
                        ReDim Preserve seqnum(seqCnt) As Integer
                        ReDim Preserve seqPin(seqCnt) As String
                        seqnum(seqCnt) = nonPowerSeqNum(UCase(pwrGloName))
                        seqPin(seqCnt) = ioPin(PinIdx)
                        seqCnt = seqCnt + 1
                    End If
                End If
            Next PinIdx
            If ioInitIdx = 0 Then
                If seqCnt > 0 Then
                    pwrUpSeq.IOHPin = seqPin
                    pwrUpSeq.IOHSeq = seqnum
                End If
            ElseIf ioInitIdx = 1 Then
                If seqCnt > 0 Then
                    pwrUpSeq.IOLPin = seqPin
                    pwrUpSeq.IOLSeq = seqnum
                End If
            Else
                If seqCnt > 0 Then
                    pwrUpSeq.IOHZPin = seqPin
                    pwrUpSeq.IOHZSeq = seqnum
                End If
            End If
        Next ioInitIdx
        ' power down
        For ioInitIdx = 0 To UBound(ioInit)
            seqCnt = 0
            For PinIdx = 0 To ioCnt - 1
                If theexec.DataManager.ChannelType(ioPin(PinIdx)) <> "N/C" Then
                    pwrGloName = ioPin(PinIdx) + PowerDownSeqName + ioInit(ioInitIdx)
                    If nonPowerSeqNum.Exists(UCase(pwrGloName)) Then
                        ReDim Preserve seqnum(seqCnt) As Integer
                        ReDim Preserve seqPin(seqCnt) As String
                        seqnum(seqCnt) = nonPowerSeqNum(UCase(pwrGloName))
                        seqPin(seqCnt) = ioPin(PinIdx)
                        seqCnt = seqCnt + 1
                        havePwrDown = True
                    End If
                End If
            Next PinIdx
            If ioInitIdx = 0 Then
                If seqCnt > 0 Then
                    pwrDownSeq.IOHPin = seqPin
                    pwrDownSeq.IOHSeq = seqnum
                End If
            ElseIf ioInitIdx = 1 Then
                If seqCnt > 0 Then
                    pwrDownSeq.IOLPin = seqPin
                    pwrDownSeq.IOLSeq = seqnum
                End If
            Else
                If seqCnt > 0 Then
                    pwrDownSeq.IOHZPin = seqPin
                    pwrDownSeq.IOHZSeq = seqnum
                End If
            End If
        Next ioInitIdx
        
        If Not havePwrDown Then
            pwrDownSeq.IOHPin = pwrUpSeq.IOLPin
            pwrDownSeq.IOLPin = pwrUpSeq.IOHPin
            pwrDownSeq.IOHZPin = pwrUpSeq.IOHZPin
            If Len(Join(pwrDownSeq.IOHPin)) > 0 Then
                ReDim pwrDownSeq.IOHSeq(UBound(pwrDownSeq.IOHPin)) As Integer
                For i = 0 To UBound(pwrDownSeq.IOHPin)
                    pwrDownSeq.IOHSeq(i) = maxSeqNum - pwrUpSeq.IOLSeq(i) + 1
                Next i
            End If
            If Len(Join(pwrDownSeq.IOLPin)) > 0 Then
                ReDim pwrDownSeq.IOLSeq(UBound(pwrDownSeq.IOLPin)) As Integer
                For i = 0 To UBound(pwrDownSeq.IOLPin)
                    pwrDownSeq.IOLSeq(i) = maxSeqNum - pwrUpSeq.IOHSeq(i) + 1
                Next i
            End If
            If Len(Join(pwrDownSeq.IOHZPin)) > 0 Then
                ReDim pwrDownSeq.IOHZSeq(UBound(pwrDownSeq.IOHZPin)) As Integer
                For i = 0 To UBound(pwrDownSeq.IOHZPin)
                    pwrDownSeq.IOHZSeq(i) = maxSeqNum - pwrUpSeq.IOHZSeq(i) + 1
                Next i
            End If
        End If
    End If
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20240311][T-All][Clyde] Add for Power Up and Power Down used, setup nWire sequence number
Public Function PowerUpDownClockSeq(seq As PowerSeqeuce, isUp As Boolean)
On Error GoTo errHandler
    Dim nWirePort() As String
    Dim nWireCnt As Integer
    Dim i As Integer
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim tmpPowerDown As String
    
    functionName = "PowerUpDownClockSeq"
    nWirePort = Split(nWire_Ports_GLB, ",")
    nWireCnt = UBound(nWirePort)
    
    ReDim seq.nWirePort(nWireCnt) As String
    ReDim seq.nWireSeq(nWireCnt) As Integer
    
    For i = 0 To nWireCnt
        Get_nWire_Name nWirePort(i), port_pa, ac_spec_pa, pin_pa, global_spec_pa
        seq.nWirePort(i) = port_pa
        
        If isUp Then
            seq.nWireSeq(i) = theexec.Specs.Globals(global_spec_pa).ContextValue
        Else
            tmpPowerDown = Replace(global_spec_pa, PowerUpSeqName, PowerDownSeqName)
            If nonPowerSeqNum.Exists(UCase(tmpPowerDown)) Then
                seq.nWireSeq(i) = theexec.Specs.Globals(tmpPowerDown).ContextValue
            Else
                seq.nWireSeq(i) = maxSeqNum - theexec.Specs.Globals(global_spec_pa).ContextValue + 1
            End If
        End If
        
    Next i
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20240311][T-All][Clyde] Add for Power Up and Power Down used, setup nWire sequence number
Public Function PowerUpDownPinSeq(Pins As String, seq As PowerSeqeuce, isUp As Boolean)
On Error GoTo errHandler
    Dim PwrPin As Variant
    Dim pin() As String
    Dim PinCnt As Long
    Dim pwrPinIndex As Long
    Dim pwrSeq As Integer
    Dim VMain As Double
    Dim Irange As Double
    
    functionName = "PowerUpDownPinSeq"
    theexec.DataManager.DecomposePinList Pins, pin, PinCnt
    
    ReDim seq.PowerSeqPin(maxSeqNum) As String
    For Each PwrPin In pin
        If gl_dicPowerPinIndex.Exists(LCase(PwrPin)) Then       'Remove N/C pin.
            pwrPinIndex = gl_dicPowerPinIndex(LCase(PwrPin))
            If isUp Then
                pwrSeq = CInt(PowerPin_range_ary(pwrPinIndex).PowerSeq)
            Else
                pwrSeq = CInt(PowerPin_range_ary(pwrPinIndex).PowerDownSeq)
            End If
            
            If pwrSeq <> 99 Then
                seq.PowerSeqPin(pwrSeq) = CombineStringList(seq.PowerSeqPin(pwrSeq), CStr(PwrPin))
            Else
                If PowerPin_range_ary(pwrPinIndex).MergeType Like "*DCVI*" Then
                    VMain = TheHdw.DCVI.Pins(PwrPin).Voltage.value
                    Irange = TheHdw.DCVI.Pins(PwrPin).CurrentRange.value
                ElseIf PowerPin_range_ary(pwrPinIndex).MergeType Like "*DCVS*" Then
                    VMain = TheHdw.DCVS.Pins(PwrPin).Voltage.value
                    Irange = TheHdw.DCVS.Pins(PwrPin).CurrentRange.value
                Else
                    Call Print_Error_Message(Error_Info, moduleName, functionName, "IO pin cannot do power up/down")
                End If
                seq.PowerSeq99Log = CombineStringList(seq.PowerSeq99Log, PowerUpDownPrintDatalogFormat(CStr(PwrPin) & "(N/A)", VMain, Irange, 0, 0, pwrSeq, isUp), vbCrLf)
            End If
        Else
            VMain = 0
            Irange = 0
            seq.PowerSeqNCLog = CombineStringList(seq.PowerSeqNCLog, PowerUpDownPrintDatalogFormat(CStr(PwrPin) & "(N/C)", VMain, Irange, 0, 0, pwrSeq, isUp), vbCrLf)
        End If
    Next PwrPin
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20240311][T-All][Clyde] Create Print Datalog format function
Public Function PowerUpDownPrintDatalogFormat(PwrPin As String, VMain As Double, Irange As Double, step As Integer, RiseTime As Double, pwrSequence As Integer, isUp As Boolean) As String
On Error GoTo errHandler
    Dim Output As String
    Dim riseFall As String
    
    functionName = "PowerUpDownPrintDatalogFormat"
    
    If isUp Then
        riseFall = "RiseTime"
    Else
        riseFall = "FallTime"
    End If
    
    Output = "print: Pin " & _
            FormatNumericDatalog(PwrPin, 30, False) & ", Vmain " & Format(VMain, "0.000") & " V, Irange " & _
            FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & _
            FormatNumericDatalog(step, 2, True) & ", " & riseFall & " " & _
            FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & _
            FormatNumericDatalog(pwrSequence, 3, True)
            
    PowerUpDownPrintDatalogFormat = Output
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20240312][All][Clyde] Modulize for FRC setup in Power up/down
Public Function SetUpPowerUpDownFRCState(nWirePort As String, seqnum As Integer, isUp As Boolean, debugF As Boolean)
On Error GoTo errHandler
    Dim printUpDown As String
    Dim nwirePortPlist As New PinList
    
    functionName = "SetUpPowerUpDownFRCState"
    
    If isUp Then
        printUpDown = "power up"
    Else
        printUpDown = "power down"
    End If
    
    theexec.Datalog.WriteComment vbCrLf & "print: " & printUpDown & " for nwire(" & seqnum & ")" & vbCrLf & RepeatChr("*", 120)
    
    If isUp Then
        nwirePortPlist.value = nWirePort
        PowerUp_Interpose nwirePortPlist, debugF
    Else
        PowerDown_Interpose nWirePort, debugF
    End If
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20240312][All][Clyde] Modulize for IO setup in Power up/down
Public Function SetUpPowerUpDownIOState(ioPins As String, seqnum As Integer, initState As ChInitState, isUp As Boolean)
On Error GoTo errHandler
    Dim Output As String
    Dim printUpDown As String
    
    functionName = "SetUpPowerUpDownIOState"
    
    TheHdw.Digital.Pins(ioPins).Connect
    TheHdw.Digital.Pins(ioPins).initState = initState
    
    If initState = chInitHi Then
        Output = "H"
    ElseIf initState = chInitLo Then
        Output = "L"
    ElseIf initState = chInitoff Then
        Output = "HZ"
    End If
    
    If isUp Then
        printUpDown = "power up"
    Else
        printUpDown = "power down"
    End If
    
    theexec.Datalog.WriteComment vbCrLf & "print: " & printUpDown & " for I/O pins force " & Output & " : (" & ioPins & ") ; sequence :" & seqnum & vbCrLf & RepeatChr("*", 120)
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function


''20240124: Added for user defined argument
Public Function CheckInstForUserFunction(UserFunction As String, digSrcLabel() As String, digSrcPatterns() As String, strPattern As String)
On Error GoTo errHandler
Dim funcName As String:: funcName = "CheckInstForUserFunction"
Dim functionArray() As String, strFunction As Variant
Dim PatArray() As String, PatCount As Long, Pat As Variant

''Initial
Dim digSrcCnt As Long: digSrcCnt = 0
ReDim digSrcLabel(0)
ReDim digSrcPatterns(0)

If UserFunction = "" Then Exit Function

''Save MFSTP pattern into
''Save MFSTP pattern into
If InStr(strPattern, ",") = 0 Then
        PatArray = TheExec.DataManager.Raw.GetPatternsInSet(strPattern, PatCount)
Else
        PatArray = Split(strPattern, ",")
End If

Dim matchedPatCount As Long: matchedPatCount = 0
For Each Pat In PatArray
    If DigSrcPatternDict.Exists(Pat) Then
        ReDim Preserve digSrcPatterns(matchedPatCount)
        digSrcPatterns(matchedPatCount) = Pat
        matchedPatCount = matchedPatCount + 1
    End If
Next Pat

''Trim and Ucase, split "DigSrc:X3G1, DigSrc:X3G2" -> "DIGSRC:X3G1" & "DIGSRC:X3G2"
functionArray = Split(UCase(Replace(UserFunction, " ", "")), ",")
For Each strFunction In functionArray
    Dim splitFunction() As String, functionType As String, functionLabel As String
    ''Split "DIGSRC:X3G1"
    splitFunction = Split(strFunction, ":")
    If UBound(splitFunction) < 1 Then Call Print_Error_Message(Error_Info, "LIB_Common", funcName, "Argument UserFunction:" & UserFunction & " has incorrect field:" & strFunction)
    
    functionType = splitFunction(0)
    functionLabel = splitFunction(1)
    
    If functionType = "DIGSRC" Then
        ReDim Preserve digSrcLabel(digSrcCnt)
        digSrcLabel(digSrcCnt) = functionLabel
        digSrcCnt = digSrcCnt + 1
    Else
        Call Print_Error_Message(Error_Info, "LIB_Common", funcName, "Argument UserFunction:" & UserFunction & " has incorrect field:" & strFunction)
    End If
Next strFunction

'Check the amount of MFSTP pattern and argument sets
If UBound(digSrcPatterns) <> UBound(digSrcLabel) Or (digSrcCnt <> matchedPatCount) Then
    Call Print_Error_Message(Error_Info, "LIB_Common", funcName, "Argument UserFunction:" & UserFunction & " does not have the same amount of MFSTP set and MFSTP pattern:" & strFunction)
End If

Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function PatternExecution(patset() As String, ReportResult As PFType, TL_C_YES As Long, ResultMode As tlResultMode, _
            ConcurrentMode As tlPatConcurrentMode, ByRef SCAN_Site_Blooean, Optional ApplyVoltageFromBinCut As String = vbNullString)

    On Error GoTo errHandler
    Dim Pat As Variant
    Dim inst_info As Instance_Info
'    Dim instrumentUtility As New Instrument_Utility
    Dim i As Long
    For Each Pat In patset
        If ApplyVoltageFromBinCut <> "" Then
           'T-Col TTR purpose for the scenario w/o selsrm pattern, 20230531
'           If LCase(Pat) Like "*_pl??_*" Then TheHdw.DCVS.Pins(Join(instrumentUtility.GetDCVSPinsFromCorePower, ",")).Voltage.Output = tlDCVSVoltageAlt

           For i = 0 To UBound(selsramLogicPingroup)
               If UCase(selsramLogicPingroup(i)) <> "PRESERVED" And UCase(selsramLogicPingroup(i)) <> "RESERVED" Then
                   If (TheHdw.DCVS.Pins(selsramLogicPingroup(i)).Voltage.Output = tlDCVSVoltageAlt) Then
                       inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                       Exit For
                   End If
               End If
           Next i

           'SycnUp
           If Flag_SyncUp_DCVS_Output_enable Then
               Call SyncUp_DCVS_Output(inst_info.p_mode, inst_info.currentDcvsOutput, SyncUp_PowerPin_Group) '''This is to sync up logic powers and sram powers on the same DCVS output (for TD testing)
           End If
        End If
        If theexec.TesterMode = testModeOffline Then
            Call ATPG_offline(CStr(Pat), ResultMode)
        Else
            If gl_bTTRDisableAlarm = False Then     'T-Col TTR approve by Si -- 230413
                TheHdw.Alarms.Check
            End If
            
            If glb_isSFC_Enabled And glb_SFC_Scan_Check = True Then
                Call Harvest_CMEM_InitSetup
            End If
            
            Call TheHdw.Patterns(CStr(Pat)).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)

        End If
        
        ''20240223: Added to avoid site boolean logical and
        Dim sitePatPass As New SiteBoolean
        sitePatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        
        Dim site As Variant
        For Each site In theexec.sites
            If sitePatPass(site) = logicFalse Then
                SCAN_Site_Blooean(site) = logicFalse
            End If
        Next site
        
        If glb_isSFC_Enabled And glb_SFC_Scan_Check = True Then
                Call SFC_CMEM_Stop
        End If
        
        'SCAN_Site_Blooean = SCAN_Site_Blooean.LogicalAnd(TheHdw.Digital.Patgen.PatternBurstPassedPerSite)

        '230711 swtich to Valt after SC Selsram pattern
'        If PrintVolatgeOutput And (UCase(CStr(Pat)) Like "*SC*") And (UCase(CStr(Pat)) Like "*_SRMDSSC*") Then
'            If TheHdw.DCVS.Pins("VDD_SOC_S1").Voltage.Output = tlDCVSVoltageMain Then
'                TheHdw.DCVS.Pins(Join(instrumentUtility.GetDCVSPinsFromCorePower, ",")).Voltage.Output = tlDCVSVoltageAlt
'                TheExec.Datalog.WriteComment "Switch to Valt after selsram pattern (by VBT)"
'                IsSwitch2Valt = True
'            Else
'                TheExec.Datalog.WriteComment "Selsram pattern switch to Valt"
'            End If
'        End If
    Next Pat

    Exit Function

errHandler:
    theexec.Datalog.WriteComment "error in PatternExecution"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20240731 Check program name
Public Function Check_Program_Name()
On Error GoTo errHandler
    Dim regex As New RegExp
    Dim illegalFormat As Boolean
    Dim c4_value As String: c4_value = vbNullString

    illegalFormat = False

    With regex
        .Global = True
        .Pattern = "^[a-zA-Z0-9\_\-\.]+$"  'Define allow character.
        .IgnoreCase = True
        illegalFormat = Not (.Test(TheExec.TestProgram.Name))
    End With

    If illegalFormat Then
        Worksheets("PinMap").Select
        c4_value = Cells(4, "C")
        Cells(4, "C") = "Program name illegal! (PinMap C4 value: " & c4_value & " )"
    Else
        'Program naming is legal.
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common", "Check_Program_Name")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ProfileData_Filter(pin As Variant, Optional InRawData As DSPWave, Optional FilterPercent As Double, Optional ForceV_Value As Double)
    
    Dim SortInWf As New DSPWave
    Dim Filter_OutputWf As New DSPWave
    Dim Filter_MAX As New DSPWave
    Dim Filter_MIN As New DSPWave
    Dim SD_Mean As New SiteDouble
    Dim SD_STDEV As New SiteDouble
    Dim SD_FilterPercent As New SiteDouble
    Dim site As Variant
    SD_FilterPercent = FilterPercent
    Filter_MAX.CreateConstant 0, 1, DspDouble
    Filter_MIN.CreateConstant 0, 1, DspDouble

    
    Call rundsp.DSP_DataFilter(InRawData, SD_FilterPercent, Filter_OutputWf, Filter_MAX, Filter_MIN, SD_Mean, SD_STDEV)
    
    THEEXEC.Flow.TestLimit resultVal:=Filter_MAX.Element(0), Tname:="ProfileData_Filter-MAX", PinName:=pin, ForceResults:=tlForceNone, ForceVal:=ForceV_Value
    THEEXEC.Flow.TestLimit resultVal:=Filter_MIN.Element(0), Tname:="ProfileData_Filter-MIN", PinName:=pin, ForceResults:=tlForceNone, ForceVal:=ForceV_Value
    THEEXEC.Flow.TestLimit resultVal:=SD_Mean, Tname:="ProfileData_Filter-AVG", PinName:=pin, ForceResults:=tlForceNone, ForceVal:=ForceV_Value
    THEEXEC.Flow.TestLimit resultVal:=SD_STDEV, Tname:="ProfileData_Filter-STDEV", PinName:=pin, ForceResults:=tlForceNone, ForceVal:=ForceV_Value
End Function
Public Function Export_All()

    Dim filename As String
    Dim Current_Dirctory As String
    Dim Folder_Bk As String
    Dim Arr_TP_Name() As String
    Dim TP_Name As String
    
    Arr_TP_Name = Split(Replace(theexec.TestProgram.PathAndName, CurDir & "\", ""), "_")
    TP_Name = Arr_TP_Name(0) & "_" & Arr_TP_Name(1)
    Folder_Bk = "D:\Local_TP\Profile_" & TP_Name
        If create_folderName = False Then
            Profile_Folder = "X" & CStr(XCoord(0)) & "Y" & CStr(YCoord(0)) & "_" & right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2) & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2)
            create_folderName = True
        End If

        If UCase(CurDir) Like "*X:\*" Then
        Call Print_Error_Message(Warning_Info, LIB_Common, "TP is in X:\, change Profile to D:\Local_TP\Profile_" & TP_Name)
        If Dir("X:\Local_TP", vbDirectory) = Empty Then
            Call Print_Error_Message(Warning_Info, LIB_Common, "D:\Local_TP folder is not exist skip export profile")
            
            Exit Function
        Else
            If Dir(Folder_Bk, vbDirectory) = Empty Then
                MkDir Folder_Bk
            Else
            End If
        End If
        
        Current_Dirctory = Folder_Bk & "\" & Profile_Folder
        Else
            Current_Dirctory = CurDir & "\" & Profile_Folder
        End If
        If Dir(Current_Dirctory, vbDirectory) = Empty Then
            MkDir Current_Dirctory
        End If

        filename = Current_Dirctory & "\" & "CurrentProfile.txt"

        Open filename For Output As #13

    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, LIB_Common, "Get_CurrentProfile_Export_All")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function ProfileRecord(Action As String)
    Dim ActionTime As Double
    Dim TempArr() As String
    Dim i As Integer
    
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Or Profile_byflow Then
        glb_TestInstance = TheExec.DataManager.instancename
    

        If Action = "Instance_Start" Then
            profile_count = 0
            T_ProfileStart = Timer()
            profileAction.RemoveAll
            profileAction.Add "ProfileRecord - " & "Instance:" & glb_TestInstance & ",Action_" & profile_count & ":" & Action & ",Time:" & Timer() - T_ProfileStart, 0
            profile_count = profile_count + 1
                   
        ElseIf Action = "plot_start" Then
            profile_count = 0
            T_ProfileStart = Timer()
            profileAction.RemoveAll
            profileAction.Add "ProfileRecord - " & "Instance:" & glb_TestInstance & ",Action_" & profile_count & ":" & Action & ",Time:" & Timer() - T_ProfileStart, 0
            profile_count = profile_count + 1
        Else
            ActionTime = Timer() - T_ProfileStart
            If Action Like "*:*" Then
                TempArr = Split(Action, ":")
                
                profileAction.Add "ProfileRecord - " & "Instance:" & glb_TestInstance & ",Action_" & profile_count & ":" & TempArr(UBound(TempArr)) & ",Time:" & Timer() - T_ProfileStart, 0
                profile_count = profile_count + 1
            ElseIf Action = "Action_Plot" Then
                
                For i = 0 To profileAction.Count - 1
                    Print #13, profileAction.Keys(i)
                    
                Next
            Else
                
                profileAction.Add "ProfileRecord - " & "Instance:" & glb_TestInstance & ",Action_" & profile_count & ":" & Action & ",Time:" & Timer() - T_ProfileStart, 0
                profile_count = profile_count + 1
            End If
        End If
    Else
    End If
End Function
