Attribute VB_Name = "LIB_HardIP_AP"
Option Explicit 'Add ErrHandler 2023/05/29
Public glb_BCCallHIPinst_passBinFromStep As New SiteLong        'Update for SelSRAM --20220613
Public glb_BCCallHIPinst_instinfo As Instance_Info              'Update for SelSRAM --20220613
Public glb_BCCallHIPinst_OverlayName As String                  'Update for SelSRAM --20220613
Public glb_SelSram_Dic As New Scripting.Dictionary              'Update for SelSRAM --20220613

Public Function Hex2Base64(ByVal sHex)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    ' Convert Hex to Base64
    Static oNode As Object
    Dim a() As Byte
    If Len(sHex) Mod 2 <> 0 Then
        sHex = left(sHex, Len(sHex) - 1) & "0" & right(sHex, 1)
    End If
    If oNode Is Nothing Then
        Set oNode = CreateObject("MSXML2.DOMDocument").CreateElement("Node")
    End If
    With oNode
        .Text = vbNullString
        .DataType = "bin.hex"
        .Text = sHex
        a = .nodeTypedValue
        .DataType = "bin.base64"
        .nodeTypedValue = a
        Hex2Base64 = .Text
    End With
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "Hex2Base64") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Server_Connection(data As String, ecid As String, chipid As String, brid As String, prod_mode As String, secure_mode As String, secure_domain As String, sitev As Variant) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim objHttp As Object
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    secure_domain = 1
    secure_mode = True
    Dim sXml As String
    Dim base64send As String
    sXml = sXml & "<?xml version=""1.0"" encoding=""UTF-8""?>"
    sXml = sXml & "<!DOCTYPE plist PUBLIC ""-//Apple//DTD PLIST 1.0//EN"" ""http://www.apple.com/DTDs/PropertyList-1.0.dtd"">"
    sXml = sXml & "<plist version=""1.0"">"
    sXml = sXml & "<dict>"
    sXml = sXml & "<key>@AuthDbg,Ticket</key>"
    sXml = sXml & "<true/>"
    sXml = sXml & "<key>AuthDbg,BoardID</key>"
    sXml = sXml & "<integer>" + brid + "</integer>"
    sXml = sXml & "<key>AuthDbg,ChipID</key>"
    sXml = sXml & "<integer>" + chipid + "</integer>"
    'sXml = sXml & "<integer>33025</integer>"
    sXml = sXml & "<key>AuthDbg,ECID</key>"
    sXml = sXml & "<integer>" + ecid + "</integer>"
    'sXml = sXml & "<integer>20015998348237</integer>"
    sXml = sXml & "<key>AuthDbg,Nonce</key>"
    sXml = sXml & "<data>"
    sXml = sXml & data
    'sXml = sXml & "McIqFULhQA4fQw4edOULM5FHqf6Duq5smgbZ78W5FX4="
    sXml = sXml & "</data>"
    sXml = sXml & "<key>AuthDbg,ProductionMode</key>"
    sXml = sXml & "<" + prod_mode + "/>"
    sXml = sXml & "<key>AuthDbg,SecurityDomain</key>"
    sXml = sXml & "<integer>" + secure_domain + "</integer>"
    sXml = sXml & "<key>AuthDbg,SecurityMode</key>"
    sXml = sXml & "<" + secure_mode + "/>"
    sXml = sXml & "<key>AuthDbg,EnAppleDebugExternal</key>"
    sXml = sXml & "<true/>"
    sXml = sXml & "<key>AuthDbg,TicketIdentifier</key>"
    sXml = sXml & "<integer>1</integer>"
    sXml = sXml & "</dict>"
    sXml = sXml & "</plist>"
    Dim url As String
    url = "https://gs.apple.com/TSS/controller?action=2"

    objHttp.Open "POST", url, False
    objHttp.setRequestHeader "Content-Type", "application/xml"
    objHttp.setoption 2, objHttp.getoption(2)
    objHttp.Send (sXml)
    Dim ResponseTxt As String
    ResponseTxt = objHttp.ResponseText

    If ResponseTxt Like "*This device isn't eligible*" Then
        TheExec.Datalog.WriteComment "<Error> Site : " & sitev & " ********************** This device isn't eligible for the requested build. **********************"
        TheExec.sites.item(sitev).FlagState("F_Auth_Device_No_Authorization") = logicTrue
        Exit Function
    End If
    
    'S = InStr(0, ResponseTxt, "<data>")
    'e = InStr(InStr(0, ResponseTxt, "<data>") + 1, ResponseTxt, "<data>")
    'len = InStr(InStr(0, ResponseTxt, "<data>") + 1, ResponseTxt, "<data>") - InStr(0, ResponseTxt, "<data>")
    Server_Connection = mid(ResponseTxt, InStr(ResponseTxt, "<data>") + 6, InStr(ResponseTxt, "</data>") - InStr(ResponseTxt, "<data>") - 6)
    Server_Connection = Replace(Server_Connection, vbTab, vbNullString)
    Server_Connection = Replace(Server_Connection, " ", vbNullString)
    Server_Connection = Replace(Server_Connection, vbLf, vbNullString)
    Server_Connection = Replace(Server_Connection, vbCrLf, vbNullString)
    
    TheExec.Datalog.WriteComment "Data from server (base64) = " & Server_Connection
    Server_Connection = Base64To16(Server_Connection)
    
    '==================================================2020/11/10_test=================================================================================
'    For i = 1 To 10
'
'        If objHttp.Status <> 200 Then
'            TheHdw.Wait 0.01
'        Else
'            objHttp.send (sXml)
'            Dim ResponseTxt As String
'            ResponseTxt = objHttp.ResponseText
'            'S = InStr(0, ResponseTxt, "<data>")
'            'e = InStr(InStr(0, ResponseTxt, "<data>") + 1, ResponseTxt, "<data>")
'            'len = InStr(InStr(0, ResponseTxt, "<data>") + 1, ResponseTxt, "<data>") - InStr(0, ResponseTxt, "<data>")
'            Server_Connection = Mid(ResponseTxt, InStr(ResponseTxt, "<data>") + 6, InStr(ResponseTxt, "</data>") - InStr(ResponseTxt, "<data>") - 6)
'            Server_Connection = Replace(Server_Connection, vbTab, "")
'            Server_Connection = Replace(Server_Connection, " ", "")
'            Server_Connection = Replace(Server_Connection, vbLf, "")
'            Server_Connection = Replace(Server_Connection, vbCrLf, "")
'
'            TheExec.Datalog.WriteComment "Data from server (base64) = " & Server_Connection
'            Server_Connection = Base64To16(Server_Connection)
'        End If
'
'    Next i
    '==================================================================================================================================================
    
Exit Function

errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "Server_Connection") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Function Base64To16(Base64 As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
  Dim Base2 As String
  Dim i As Long
  For i = 1 To Len(Base64)
    Select Case mid(Base64, i, 1)
      Case "A"
        Base2 = Base2 & "000000"
      Case "B"
        Base2 = Base2 & "000001"
      Case "C"
        Base2 = Base2 & "000010"
      Case "D"
        Base2 = Base2 & "000011"
      Case "E"
        Base2 = Base2 & "000100"
      Case "F"
        Base2 = Base2 & "000101"
      Case "G"
        Base2 = Base2 & "000110"
      Case "H"
        Base2 = Base2 & "000111"
      Case "I"
        Base2 = Base2 & "001000"
      Case "J"
        Base2 = Base2 & "001001"
      Case "K"
        Base2 = Base2 & "001010"
      Case "L"
        Base2 = Base2 & "001011"
      Case "M"
        Base2 = Base2 & "001100"
      Case "N"
        Base2 = Base2 & "001101"
      Case "O"
        Base2 = Base2 & "001110"
      Case "P"
        Base2 = Base2 & "001111"
      Case "Q"
        Base2 = Base2 & "010000"
      Case "R"
        Base2 = Base2 & "010001"
      Case "S"
        Base2 = Base2 & "010010"
      Case "T"
        Base2 = Base2 & "010011"
      Case "U"
        Base2 = Base2 & "010100"
      Case "V"
        Base2 = Base2 & "010101"
      Case "W"
        Base2 = Base2 & "010110"
      Case "X"
        Base2 = Base2 & "010111"
      Case "Y"
        Base2 = Base2 & "011000"
      Case "Z"
        Base2 = Base2 & "011001"
      Case "a"
        Base2 = Base2 & "011010"
      Case "b"
        Base2 = Base2 & "011011"
      Case "c"
        Base2 = Base2 & "011100"
      Case "d"
        Base2 = Base2 & "011101"
      Case "e"
        Base2 = Base2 & "011110"
      Case "f"
        Base2 = Base2 & "011111"
      Case "g"
        Base2 = Base2 & "100000"
      Case "h"
        Base2 = Base2 & "100001"
      Case "i"
        Base2 = Base2 & "100010"
      Case "j"
        Base2 = Base2 & "100011"
      Case "k"
        Base2 = Base2 & "100100"
      Case "l"
        Base2 = Base2 & "100101"
      Case "m"
        Base2 = Base2 & "100110"
      Case "n"
        Base2 = Base2 & "100111"
      Case "o"
        Base2 = Base2 & "101000"
      Case "p"
        Base2 = Base2 & "101001"
      Case "q"
        Base2 = Base2 & "101010"
      Case "r"
        Base2 = Base2 & "101011"
      Case "s"
        Base2 = Base2 & "101100"
      Case "t"
        Base2 = Base2 & "101101"
      Case "u"
        Base2 = Base2 & "101110"
      Case "v"
        Base2 = Base2 & "101111"
      Case "w"
        Base2 = Base2 & "110000"
      Case "x"
        Base2 = Base2 & "110001"
      Case "y"
        Base2 = Base2 & "110010"
      Case "z"
        Base2 = Base2 & "110011"
      Case "0"
        Base2 = Base2 & "110100"
      Case "1"
        Base2 = Base2 & "110101"
      Case "2"
        Base2 = Base2 & "110110"
      Case "3"
        Base2 = Base2 & "110111"
      Case "4"
        Base2 = Base2 & "111000"
      Case "5"
        Base2 = Base2 & "111001"
      Case "6"
        Base2 = Base2 & "111010"
      Case "7"
        Base2 = Base2 & "111011"
      Case "8"
        Base2 = Base2 & "111100"
      Case "9"
        Base2 = Base2 & "111101"
      Case "+"
        Base2 = Base2 & "111110"
      Case "/"
        Base2 = Base2 & "111111"
      Case "="
        Base2 = Base2 & "000000"
      Case Else
       'Debug.Print Mid(Base64, i, 1)
        'Base64To16 = CVErr(xlErrValue)
        Exit Function
    End Select
  Next i
  If Not Len(Base2) Mod 4 = 0 Then
    Base2 = String(4 - (Len(Base2) Mod 4), "0") & Base2
  End If
  If Len(Base2) > 4 And left(Base2, 4) = "0000" Then
    Base2 = mid(Base2, 5)
  End If
  Base64To16 = vbNullString
  For i = 1 To Len(Base2) Step 4
    Select Case mid(Base2, i, 4)
      Case "0000"
        Base64To16 = Base64To16 & "0"
      Case "0001"
        Base64To16 = Base64To16 & "1"
      Case "0010"
        Base64To16 = Base64To16 & "2"
      Case "0011"
        Base64To16 = Base64To16 & "3"
      Case "0100"
        Base64To16 = Base64To16 & "4"
      Case "0101"
        Base64To16 = Base64To16 & "5"
      Case "0110"
        Base64To16 = Base64To16 & "6"
      Case "0111"
        Base64To16 = Base64To16 & "7"
      Case "1000"
        Base64To16 = Base64To16 & "8"
      Case "1001"
        Base64To16 = Base64To16 & "9"
      Case "1010"
        Base64To16 = Base64To16 & "A"
      Case "1011"
        Base64To16 = Base64To16 & "B"
      Case "1100"
        Base64To16 = Base64To16 & "C"
      Case "1101"
        Base64To16 = Base64To16 & "D"
      Case "1110"
        Base64To16 = Base64To16 & "E"
      Case "1111"
        Base64To16 = Base64To16 & "F"
      Case Else
    End Select
  Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "Base64To16") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function binToDecStr(Optional default_value As String = "1111111111111111111111111111111111111111111111111111111111111110") As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    Dim table(65) As String
    Dim overflow(40) As String
    Dim final_value As String
    Dim final_str As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim temp_calc
    
    'Dim default_value As String: default_value = "1111111111111111111111111111111111111111111111111111111111111110" 'LSB -> MSB
    
    table(0) = "0000000000000000000000000000000000000001"
    table(1) = "0000000000000000000000000000000000000002"
    table(2) = "0000000000000000000000000000000000000004"
    table(3) = "0000000000000000000000000000000000000008"
    table(4) = "0000000000000000000000000000000000000016"
    table(5) = "0000000000000000000000000000000000000032"
    table(6) = "0000000000000000000000000000000000000064"
    table(7) = "0000000000000000000000000000000000000128"
    table(8) = "0000000000000000000000000000000000000256"
    table(9) = "0000000000000000000000000000000000000512"
    table(10) = "0000000000000000000000000000000000001024"
    table(11) = "0000000000000000000000000000000000002048"
    table(12) = "0000000000000000000000000000000000004096"
    table(13) = "0000000000000000000000000000000000008192"
    table(14) = "0000000000000000000000000000000000016384"
    table(15) = "0000000000000000000000000000000000032768"
    table(16) = "0000000000000000000000000000000000065536"
    table(17) = "0000000000000000000000000000000000131072"
    table(18) = "0000000000000000000000000000000000262144"
    table(19) = "0000000000000000000000000000000000524288"
    table(20) = "0000000000000000000000000000000001048576"
    table(21) = "0000000000000000000000000000000002097152"
    table(22) = "0000000000000000000000000000000004194304"
    table(23) = "0000000000000000000000000000000008388608"
    table(24) = "0000000000000000000000000000000016777216"
    table(25) = "0000000000000000000000000000000033554432"
    table(26) = "0000000000000000000000000000000067108864"
    table(27) = "0000000000000000000000000000000134217728"
    table(28) = "0000000000000000000000000000000268435456"
    table(29) = "0000000000000000000000000000000536870912"
    table(30) = "0000000000000000000000000000001073741824"
    table(31) = "0000000000000000000000000000002147483648"
    table(32) = "0000000000000000000000000000004294967296"
    table(33) = "0000000000000000000000000000008589934592"
    table(34) = "0000000000000000000000000000017179869184"
    table(35) = "0000000000000000000000000000034359738368"
    table(36) = "0000000000000000000000000000068719476736"
    table(37) = "0000000000000000000000000000137438953472"
    table(38) = "0000000000000000000000000000274877906944"
    table(39) = "0000000000000000000000000000549755813888"
    table(40) = "0000000000000000000000000001099511627776"
    table(41) = "0000000000000000000000000002199023255552"
    table(42) = "0000000000000000000000000004398046511104"
    table(43) = "0000000000000000000000000008796093022208"
    table(44) = "0000000000000000000000000017592186044416"
    table(45) = "0000000000000000000000000035184372088832"
    table(46) = "0000000000000000000000000070368744177664"
    table(47) = "0000000000000000000000000140737488355328"
    table(48) = "0000000000000000000000000281474976710656"
    table(49) = "0000000000000000000000000562949953421312"
    table(50) = "0000000000000000000000001125899906842624"
    table(51) = "0000000000000000000000002251799813685248"
    table(52) = "0000000000000000000000004503599627370496"
    table(53) = "0000000000000000000000009007199254740992"
    table(54) = "0000000000000000000000018014398509481984"
    table(55) = "0000000000000000000000036028797018963968"
    table(56) = "0000000000000000000000072057594037927936"
    table(57) = "0000000000000000000000144115188075855872"
    table(58) = "0000000000000000000000288230376151711744"
    table(59) = "0000000000000000000000576460752303423488"
    table(60) = "0000000000000000000001152921504606846976"
    table(61) = "0000000000000000000002305843009213693952"
    table(62) = "0000000000000000000004611686018427387904"
    table(63) = "0000000000000000000009223372036854775808"
    table(64) = "0000000000000000000018446744073709551616"
    
    For i = 40 To 1 Step -1
    
        temp_calc = 0
    
            For j = 1 To Len(default_value)
                If CLng(mid(default_value, j, 1)) = 1 Then
                    temp_calc = temp_calc + CLng(mid(table(j - 1), i, 1))
                End If
            Next
            
            If overflow(i) = "" Then overflow(i) = "0"
            temp_calc = temp_calc + CLng(overflow(i))
            
            final_value = mid(CStr(temp_calc), Len(temp_calc), 1) & final_value
            
            
            For k = 1 To Len(temp_calc)
                If k > 1 Then
                    If overflow(i - k + 1) = "" Then overflow(i - k + 1) = "0"
                    overflow(i - k + 1) = CStr(CLng(overflow(i - k + 1)) + CLng(mid(CStr(temp_calc), Len(CStr(temp_calc)) - k + 1, 1)))
                    If Len(overflow(i - k + 1)) > 1 Then
                        If overflow(i - k) = "" Then overflow(i - k) = "0"
                        overflow(i - k) = mid(overflow(i - k + 1), 1, 1)
                        overflow(i - k + 1) = mid(overflow(i - k + 1), 2, 1)
                    End If
                End If
            Next
    Next
    
    
    For i = 1 To Len(final_value)
    
        If mid(final_value, i, 1) <> "0" Then
        
            final_str = mid(final_value, i, Len(final_value) - 1)
            Exit For
        End If
    Next
    
    
    binToDecStr = final_str
    
    If final_str = "" Then binToDecStr = "0"

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "binToDecStr") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
'Update for SelSRAM --20220613
Public Function HardIP_Decide_Switching_Bit(Logic_PinName As String, digSrc_EQ As String, ByRef digsrc_value As SiteVariant, pl_DSSC_Switching_Voltage As PinListData, PattIdx As Long, ByRef sl_selsrm_bitArray() As SiteLong) As SiteVariant
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim p_ary() As String, p_cnt As Long
    Dim logicPin As String
    Dim SramPin As String
    Dim DSSC_Switching_Voltage As New PinListData
    Dim Sdomain() As Long
    Dim DSSCSelSrmOpposite As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim ReturnString As New SiteVariant
    Dim LogicValue() As Double
    Dim SramValue As Double
    Dim PowerDomain_Idx As Long
    

    For i = 0 To UBound(SelsramMapping(PattIdx).logic_Pin)
        If UCase(SelsramMapping(PattIdx).DigSrc_Assignment(i)) = UCase(Logic_PinName) Then
            logicPin = SelsramMapping(PattIdx).logic_Pin(i)
            SramPin = SelsramMapping(PattIdx).SRAM_PIN(i)
            DSSCSelSrmOpposite = SelsramMapping(PattIdx).SelSrm1(i)
            Call HardIP_Decide_DSSC_Switching_Voltage(pl_DSSC_Switching_Voltage, logicPin, SramPin)
            PowerDomain_Idx = VddBinStr2Enum(logicPin)
            Exit For
        End If
    Next i
    '''--------------If the DigSrc Assignment does not belong to the SelSRAM Pin.--------------
    If logicPin = "" Then
        digsrc_value = digSrc_EQ
        Exit Function
    End If
    '''--------------If the DigSrc Assignment does not belong to the SelSRAM Pin.--------------

        '[20231106][All][Neil] Update Selsram function judgement for multi-SRAM pin
    Dim s_ReturnString() As String
    Dim DSPWaveSwitch As New DSPWave
    Call SelSRAM_DigSrc_Bit(PattIdx, Logic_PinName, pl_DSSC_Switching_Voltage, DSPWaveSwitch, s_ReturnString, sl_selsrm_bitArray, True)
    digsrc_value = sl_selsrm_bitArray(PowerDomain_Idx)
    Exit Function
  
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "HardIP_Decide_Switching_Bit") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

'Update for SelSRAM --20220613
Public Function HardIP_Decide_DSSC_Switching_Voltage(DSSC_Switching_Voltage As PinListData, logic_Pin As String, SRAM_PIN As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim p_ary() As String, p_cnt As Long, i As Long
''    Dim Site As Variant
    Dim j As Long
    Dim sd_Temp As New SiteDouble
    
    If logic_Pin <> "" Then
        TheExec.DataManager.DecomposePinList logic_Pin, p_ary, p_cnt
        For j = 0 To p_cnt - 1
            If TheExec.DataManager.ChannelType(p_ary(j)) <> "N/C" Then
                If glb_SelSram_Dic.Exists(LCase(p_ary(j))) Then
                    If bBCcallHIP Then ''Fixed Value for each performance mode when BinCut Usage
                        sd_Temp = glb_BCCallHIPinst_instinfo.voltage_SelsrmBitCalc(VddBinStr2Enum(p_ary(j))) / 1000
                    Else
                        sd_Temp = FormatNumber(TheHdw.DCVS.pins(p_ary(j)).Voltage.Alt, 3)
                    End If
                    DSSC_Switching_Voltage.pins(p_ary(j)) = sd_Temp
''                    DSSC_Switching_Voltage.Pins(p_ary(j)) = TheHdw.DCVS.Pins(p_ary(j)).Voltage.Alt.value
                Else
                    If bBCcallHIP Then ''Fixed Value for each performance mode when BinCut Usage
                        sd_Temp = glb_BCCallHIPinst_instinfo.voltage_SelsrmBitCalc(VddBinStr2Enum(p_ary(j))) / 1000
                    Else
                        sd_Temp = FormatNumber(TheHdw.DCVS.pins(p_ary(j)).Voltage.Alt, 3)
                    End If
                    glb_SelSram_Dic.Add LCase(p_ary(j)), LCase(p_ary(j))
                    DSSC_Switching_Voltage.AddPin p_ary(j)
                    DSSC_Switching_Voltage.pins(p_ary(j)) = sd_Temp
''                    DSSC_Switching_Voltage.Pins(p_ary(j)) = TheHdw.DCVS.Pins(p_ary(j)).Voltage.Alt.value
                End If
            End If
        Next j
    End If
    
    If SRAM_PIN <> "" Then
        TheExec.DataManager.DecomposePinList SRAM_PIN, p_ary, p_cnt
        For j = 0 To p_cnt - 1
            If TheExec.DataManager.ChannelType(p_ary(j)) <> "N/C" Then
                If glb_SelSram_Dic.Exists(LCase(p_ary(j))) Then
                    If bBCcallHIP Then ''Fixed Value for each performance mode when BinCut Usage
                        sd_Temp = glb_BCCallHIPinst_instinfo.sram_Vth(VddBinStr2Enum(logic_Pin)) / 1000
                    Else
                        sd_Temp = FormatNumber(TheHdw.DCVS.pins(p_ary(j)).Voltage.Alt, 3)
                    End If
                    DSSC_Switching_Voltage.pins(p_ary(j)) = sd_Temp
                Else
                    If bBCcallHIP Then ''Fixed Value for each performance mode when BinCut Usage
                        sd_Temp = glb_BCCallHIPinst_instinfo.sram_Vth(VddBinStr2Enum(logic_Pin)) / 1000
                    Else
                        sd_Temp = FormatNumber(TheHdw.DCVS.pins(p_ary(j)).Voltage.Alt, 3)
                    End If
                    glb_SelSram_Dic.Add LCase(p_ary(j)), LCase(p_ary(j))
                    DSSC_Switching_Voltage.AddPin p_ary(j)
                    DSSC_Switching_Voltage.pins(p_ary(j)) = sd_Temp
                End If
            End If
        Next j
    End If

    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "HardIP_Decide_DSSC_Switching_Voltage") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
 End Function

'Update for SelSRAM --20220613
Public Function Printout_BCCallHIP_DigSrc(vsite As Variant, blockName As String, sl_selsrm_bitArray() As SiteLong)
    On Error GoTo errHandler
    
    Dim site As Variant
    Dim i As Integer
    Dim j As Long
    Dim PattIdx As Long
    Dim powerDomain As String
    Dim str_bitCompareAll As String
    Dim str_bitCompare As String
    Dim str_bitDSSC As String
    Dim str_sram_vth As String
    Dim str_logic_voltage As String
    Dim alphaPowerDomain As String
    
    Dim str_Selsrm_DSSC_Bit As String
    
    Dim str_Selsrm_srm_Pin As String
    
    Dim PowerDomain_Idx As Long

    For i = 0 To UBound(SelsramMapping)
        If UCase(SelsramMapping(i).blockName) <> "*" And blockName Like UCase(SelsramMapping(i).blockName) Then
            PattIdx = i
            Exit For
        End If
    Next i
    
    For i = 0 To UBound(SelsramMapping(PattIdx).logic_Pin)
        powerDomain = UCase(Trim(SelsramMapping(PattIdx).logic_Pin(i)))
        If powerDomain <> "PRESERVED" And powerDomain <> "RESERVED" Then
            PowerDomain_Idx = VddBinStr2Enum(powerDomain)
            If str_Selsrm_DSSC_Bit = "" Then
                str_Selsrm_DSSC_Bit = sl_selsrm_bitArray(PowerDomain_Idx)
            Else
                str_Selsrm_DSSC_Bit = str_Selsrm_DSSC_Bit & sl_selsrm_bitArray(PowerDomain_Idx)
            End If
        Else
            str_Selsrm_DSSC_Bit = str_Selsrm_DSSC_Bit & 0
        End If
    Next i
    
    If Flag_Remove_Printing_BV_voltages = False And Flag_Skip_Printing_SelSrm_DSSC_Info = False Then
''        For Each site In theexec.sites
            '''//Prefix
            str_bitCompare = "SELSRAM_Compare_Bit_Str," & vsite & ","
            str_bitCompareAll = vbNullString
            str_bitDSSC = "SELSRAM_DSSC_Bit_Str," & vsite & "," & str_Selsrm_DSSC_Bit
            str_sram_vth = "SRAM_Vth(DCVS)," & vsite
            str_logic_voltage = "SelSram_voltage," & vsite
    
            For i = 0 To UBound(SelsramMapping(PattIdx).logic_Pin)
                powerDomain = UCase(Trim(SelsramMapping(PattIdx).logic_Pin(i)))
                str_Selsrm_srm_Pin = UCase(Trim(SelsramMapping(PattIdx).SRAM_PIN(i)))
                '''//Check if powerDomain is "PRESERVED" or "RESERVED". C651 Toby requested this on 20200916.
                If powerDomain <> "PRESERVED" And powerDomain <> "RESERVED" Then
                    alphaPowerDomain = UCase(SelsramMapping(PattIdx).alpha(i))
                    
                    str_bitCompare = str_bitCompare & sl_selsrm_bitArray(VddBinStr2Enum(powerDomain))
                    str_bitCompareAll = str_bitCompareAll & "," & alphaPowerDomain & "=" & sl_selsrm_bitArray(VddBinStr2Enum(powerDomain))
                    str_sram_vth = str_sram_vth & "," & alphaPowerDomain & "=" & CStr(Format(Floor(glb_BCCallHIPinst_instinfo.sram_Vth(VddBinStr2Enum(powerDomain))(vsite)) / 1000, "#0.000")) & "V"
                    str_logic_voltage = str_logic_voltage & "," & alphaPowerDomain & "=" & CStr(Format(Floor(glb_BCCallHIPinst_instinfo.voltage_SelsrmBitCalc(VddBinStr2Enum(powerDomain))(vsite)) / 1000, "#0.000")) & "V"
                End If
            Next i
            
            TheExec.Datalog.WriteComment str_bitCompare & "(LSB->MSB)" & str_bitCompareAll & vbCrLf & str_bitDSSC & vbCrLf & str_sram_vth & vbCrLf & str_logic_voltage
''        Next site
    End If
    
    Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "Printout_BCCallHIP_DigSrc") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function TrimCodeFreq_RunPat_and_MeasF(pat As String, TestSequence As String, CPU_Flag_In_Pat As Boolean, _
    DigSrc_pin As PinList, DigSrc_SampleSize As Long, MeasureF_Pin As PinList, MeasureFreq As PinListData, _
    InDSPwave As DSPWave, Interpose_PreMeas As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    'TrimCodeFreq_RunPat_and_MeasF(Pat,TestSequence,CPU_Flag_In_Pat,DigSrc_Pin,DigSrc_SampleSize,MeasureF_Pin,MeasureFreq,InDSPwave,Interpose_PreMeas)

    Dim Ts As Variant
    Dim TestSeqNum As Long
    Dim TestSequenceArray() As String
    Dim Interpose_PreMeas_Ary() As String
    Dim d_MeasF_Interval  As Double
    Dim sSrcSigName As String

    TestSeqNum = 0
    TestSequenceArray = Split(TestSequence, ",")
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    d_MeasF_Interval = 0.001 ''20190903 0.001
    'd_MeasF_Interval = 0.01 ''20190903 0.001

    Dim tempVarArray As Variant
    tempVarArray = TheHdw.DSSC.pins(DigSrc_pin).Pattern(pat).Source.Labels.list ''20210609 temp
    sSrcSigName = tempVarArray(0)


    'Call SetupDigSrcDspWave(Pat, DigSrc_pin, "TrimCodeFreq", DigSrc_SampleSize, InDSPwave)
    Call SetupDigSrcDspWave(pat, DigSrc_pin, sSrcSigName, DigSrc_SampleSize, InDSPwave)
    Call TheHdw.patterns(pat).start

    For Each Ts In TestSequenceArray
        If (CPU_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            Else
            'Do nothing
            End If
        End If

        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
            If TheExec.TesterMode = testModeOffline Then
                Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
            End If
        End If

        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            Else
            'Do nothing
            End If
        End If
        TestSeqNum = TestSeqNum + 1
        If (CPU_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    TheHdw.Digital.Patgen.HaltWait

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "TrimCodeFreq_RunPat_and_MeasF") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function TrimCodeFreq_WriteComment_DspTrimCode(InDsp As DSPWave)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim OutputTrimCode As String
    Dim TempCnt As Long
    Dim site As Variant

    For Each site In TheExec.sites
        OutputTrimCode = vbNullString
        For TempCnt = 0 To InDsp(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(InDsp(site).Element(TempCnt))
            If TempCnt = 11 Then ''11
                OutputTrimCode = OutputTrimCode & ","
            End If
        Next TempCnt
        TheExec.Datalog.WriteComment ("Site " & site & " Output Trim Code = " & OutputTrimCode)
    Next site
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_AP", "TrimCodeFreq_WriteComment_DspTrimCode") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
