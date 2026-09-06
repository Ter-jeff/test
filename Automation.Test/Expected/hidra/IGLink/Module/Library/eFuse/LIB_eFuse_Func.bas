Attribute VB_Name = "LIB_eFuse_Func"
Option Explicit
Public Function auto_bitStr2Dec(bitStr As String, Optional bitstrM_flag As Boolean = True) As Long
On Error GoTo errHandler
Dim i As Long
Dim m_dec As Long
Dim bitwidth As Long
    ''''EX:
    ''''bitstr=11001, m_dec=25
    bitwidth = Len(bitStr)
    m_dec = 0
    
    ''''case: bitstr is [LSB...MSB]
    ''''Then set bitstrM_flag to Fasle, bitstr should be reversed to [MSB...LSB]
    If (bitstrM_flag = False) Then
        bitStr = StrReverse(bitStr)
    End If
    
    ''''<NOTICE>
    ''''if BitWidth >31, it will result in an overflow error message and supposedy it's a reserved bits.
    If (bitwidth <= 31) Then
        For i = 0 To bitwidth - 1
            m_dec = m_dec + CLng(mid(bitStr, i + 1, 1)) * (2 ^ (bitwidth - 1 - i))
        Next i
    Else
        ''supposedy it's a reserved bits and all '0'.
        m_dec = 0
    End If

    auto_bitStr2Dec = m_dec
    
    ''Debug.Print "bitstr=" + bitstr + ", Dec=" + CStr(m_dec)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_bitStr2Dec")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_Dec2Bin_EFuse(ByVal n As Long, bitwidth As Long, ByRef BinArray() As Long) As String
On Error GoTo errHandler
Dim i As Long
Dim m_bitStrM As String
    ''Debug.Print "n = " & n & ", bitwidth=" & BitWidth
''''-----------------------------------
''''<Example>
''''n = 11, bitwidth=6
''''BinArray [0] = 1
''''BinArray [1] = 1
''''BinArray [2] = 0
''''BinArray [3] = 1
''''BinArray [4] = 0
''''BinArray [5] = 0
''''m_bitstrM [MSB...LSB] = 001011
''''-----------------------------------

    ''Initialize the content of array
    ReDim BinArray(bitwidth - 1) ''''BinArray[0] is LSB
    m_bitStrM = vbNullString

''''    ''''20171117 update to gate the negative input
''''    If (n < 0) Then
''''        TheExec.AddOutput "<Error> " + funcName + ":: the input n=" + CStr(n) + " is the negative value."
''''        TheExec.Datalog.WriteComment "<Error> " + funcName + ":: the input n=" + CStr(n) + " is the negative value."
''''        GoTo errHandler
''''    End If

    For i = 0 To bitwidth - 1
        BinArray(i) = 0
        If (n Mod 2) Then
            BinArray(i) = 1
        Else
            BinArray(i) = 0
        End If
        m_bitStrM = CStr(BinArray(i)) + m_bitStrM ''''[MSB...LSB]
        n = Fix(n / 2)
        ''Debug.Print "BinArray[" & i & "] = " & BinArray(i)
    Next i
    ''Debug.Print "m_bitstrM[MSB...LSB] = " & m_bitstrM
    
    auto_Dec2Bin_EFuse = m_bitStrM

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_Dec2Bin_EFuse")
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_HexStr2Value(ByVal HexStr As String) As Variant
On Error GoTo errHandler
Dim i As Long
Dim m_HexStr As String ''''without the prefix '0x' or 'x'
Dim m_len As Long
Dim m_char As String
Dim m_chVal As Double
Dim m_hex2Val As Double
    ''''-------------------------------------------------------------------------
    ''''Example, it could support up to (2^1024 -1)
    ''''-------------------------------------------------------------------------
    ''''Call auto_HexStr2Value("0x7FFFFFFF")
    ''''auto_HexStr2Value:: 0x7FFFFFFF = 2147483647
    ''''auto_HexStr2Value:: 0xFFFF = 65535
    ''''auto_HexStr2Value:: 0x2F = 47
    ''''Call auto_HexStr2Value("0x80000000")
    ''''auto_HexStr2Value:: 0x80000000 = 2147483648
    ''''Call auto_HexStr2Value("0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    ''''auto_HexStr2Value::     0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF = 1.34078079299426E+154
    ''''Call auto_HexStr2Value("0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    ''''auto_HexStr2Value::     0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF = 1.15792089237316E+77
    ''''auto_HexStr2Value::     0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF = 3.40282366920938E+38
    ''''auto_HexStr2Value::     0xFFFFFFFFFFFFFFFF = 1.84467440737096E+19
    ''''auto_HexStr2Value::     0x8FFFFFFF = 2415919103
    ''''-------------------------------------------------------------------------

    HexStr = UCase(Trim(HexStr))
    m_hex2Val = 0
    HexStr = Replace(HexStr, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE

    If (auto_isHexString(HexStr) = True) Then
        ''''20170811 update
        If (HexStr Like "0X*") Then
            m_HexStr = Replace(HexStr, "0X", "", 1)
        ElseIf (HexStr Like "X*") Then
            m_HexStr = Replace(HexStr, "X", "", 1)
        End If
        m_len = Len(m_HexStr)
        For i = 1 To m_len
            m_char = mid(m_HexStr, i, 1)
            Select Case m_char
                Case "A"
                    m_chVal = 10
                Case "B"
                    m_chVal = 11
                Case "C"
                    m_chVal = 12
                Case "D"
                    m_chVal = 13
                Case "E"
                    m_chVal = 14
                Case "F"
                    m_chVal = 15
                Case Else
                    m_chVal = CDbl(m_char)
            End Select
            ''''20180522 update if HEX characters is over 255 (>1023bits)
            If ((m_len - i) < 256) Then
                m_hex2Val = m_hex2Val + m_chVal * (16 ^ (m_len - i))
            ElseIf ((m_len - i) >= 256 And m_chVal = 0) Then
                m_hex2Val = m_hex2Val + 0
            Else
                auto_HexStr2Value = HexStr
                ''TheExec.AddOutput "<WARNING> " + funcName + ":: " + HexStr + " is over 1023bits(255 Hex_Characters)."
                ''TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: " + HexStr + " is over 1023bits(255 Hex_Characters)."
                ''Debug.Print funcName + ":: " + HexStr + " = " + CStr(auto_HexStr2Value)
                Exit Function
            End If
        Next i
        auto_HexStr2Value = m_hex2Val
        ''TheExec.Datalog.WriteComment funcName + ":: " + hexStr + " = " + CStr(m_hex2Val)
        ''Debug.Print funcName + ":: " + HexStr + " = " + CStr(m_hex2Val)
    Else
        auto_HexStr2Value = 0
        Call Print_Error_Message(Warning_Info, "LIB_eFuse_Func", "auto_HexStr2Value", "auto_HexStr2Value" + ":: " + HexStr + " is NOT a Hex String (with the prefix '0x' or 'x').")
        GoTo errHandler
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_HexStr2Value")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_isHexString(ByVal Inputstr As String) As Boolean
On Error GoTo errHandler
Dim i As Long, j As Long
Dim m_len As Long
Dim m_char As String
Dim HexChar() As Variant
Dim m_match_flag As Boolean

    Inputstr = UCase(Inputstr)
    
    HexChar = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")

    m_match_flag = False ''''<MUST> initialize
    
    Inputstr = Replace(Inputstr, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE

    ''''20170811 update
    If Inputstr Like UCase("0X*") Then ''''case "0xABCD"
        Inputstr = Replace(Inputstr, "0X", "", 1, 1)
        m_match_flag = True
    ElseIf Inputstr Like UCase("X*") Then ''''case "xABCD"
        Inputstr = Replace(Inputstr, "X", "", 1, 1)
        m_match_flag = True
    Else
        m_match_flag = False
        ''TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: It's NOT the HEX String format, should have prefix '0x' as Hex-String."
    End If

    ''''do the advanced analysis
    If (m_match_flag = True) Then
        m_len = Len(Inputstr)
        For i = 1 To m_len
            m_match_flag = False ''''<MUST> initialize per character, 20160616 update
            m_char = mid(Inputstr, i, 1)
            For j = 0 To UBound(HexChar)
                If (m_char = CStr(HexChar(j))) Then
                    m_match_flag = True
                    Exit For
                End If
            Next j
            If (m_match_flag = False) Then Exit For ''''<NOTICE>
        Next i
    End If
    auto_isHexString = m_match_flag
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_isHexString")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UpdateDLogColumns(tsNameWidth As Variant) As Long
On Error GoTo errHandler
    ''''20170217 update
    If (gB_newDlog_Flag) Then Exit Function

    tsNameWidth = CLng(tsNameWidth)
    theexec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    With theexec.Datalog.Setup.Shared.ascii.Columns.Parametric
        .TestName.Enable = True
        .TestName.Width = tsNameWidth
        .pin.Enable = True
        .pin.Width = 25
    End With
    With theexec.Datalog.Setup.Shared.ascii.Columns.Functional
        .Pattern.Enable = True
        .Pattern.Width = 128 '.Pattern.DefaultWidth
        .TestName.Enable = True
        .TestName.Width = tsNameWidth
    End With
    theexec.Datalog.ApplySetup

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "UpdateDLogColumns")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_checkIEDAString(inStr1 As String) As String
On Error GoTo errHandler
Dim k As Long
Dim tmpA() As String
Dim tmpstr As String
Dim tmpA_size As Long

    tmpA = Split(inStr1, ",")
    tmpA_size = UBound(tmpA) + 1
    
    ''Debug.Print "tmpA size =" & tmpA_size
    tmpstr = vbNullString
    If (tmpA_size < theexec.sites.Existing.Count) Then
        If (tmpA_size = 0) Then
            tmpstr = "NA"
            For k = 1 To (theexec.sites.Existing.Count - tmpA_size - 1)
                tmpstr = tmpstr + ",NA"
            Next k
        Else
            For k = 1 To (theexec.sites.Existing.Count - tmpA_size)
                tmpstr = tmpstr + ",NA"
            Next k
        End If
        theexec.Datalog.WriteComment "auto_checkIEDAString" + ":: original = " + inStr1
        inStr1 = inStr1 + tmpstr
        theexec.Datalog.WriteComment Space(23) + "  update = " + inStr1
        Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_checkIEDAString", "Could have the site sequence problem (case1)")
        theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:="IEDAString check fail"
    ElseIf (tmpA_size > theexec.sites.Existing.Count) Then ''''should not have this case
        tmpstr = vbNullString
        For k = 0 To theexec.sites.Existing.Count - 1
            If (k = (theexec.sites.Existing.Count - 1)) Then
                tmpstr = tmpstr + tmpA(k)
            Else
                tmpstr = tmpstr + tmpA(k) + ","
            End If
        Next k
        theexec.Datalog.WriteComment "auto_checkIEDAString" + ":: original = " + inStr1
        inStr1 = tmpstr
        theexec.Datalog.WriteComment Space(23) + "  update = " + inStr1
        Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_checkIEDAString", "Could have the site sequence problem (case2)")
        theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:="IEDA check fail"
    End If
    
    auto_checkIEDAString = inStr1
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_checkIEDAString")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FormatNumeric(num As Variant, length As Long) As String
On Error GoTo errHandler
Dim numStr As String
Dim tmpLen As Long
Dim spcLen As Long
    ''''Example
    ''''----------------------------------------
    '''' length > 0  is to right shift
    '''' length < 0  is to left  shift
    ''''----------------------------------------
    ''''FormatNumeric(123456, 8) + "...end"
    ''''  123456...end
    ''''
    ''''FormatNumeric(123456,-8) + "...end"
    ''''123456  ...end
    ''''
    ''''----------------------------------------
    numStr = CStr(num)
    tmpLen = Len(numStr)
    
    If (tmpLen > Abs(length)) Then
        spcLen = 0
    Else
        spcLen = Abs(length) - tmpLen
    End If
    
    If (length < 0) Then   ''''number shift to the very left
        FormatNumeric = numStr + Space(spcLen)
    Else ''''default: shift to the very right
        FormatNumeric = Space(spcLen) + numStr
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "FormatNumeric")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_eFuse_GetCatenameMaxLen(ByVal FuseType As String, Optional showPrint As Boolean = False) As Long
On Error GoTo errHandler
Dim m_len As Long
Dim m_dlogstr As String
Dim gI_ECID_catename_maxLen As Variant, gI_CFG_catename_maxLen As Variant, gI_UID_catename_maxLen As Variant, gI_UDR_catename_maxLen As Variant, gI_SEN_catename_maxLen As Variant
Dim gI_MON_catename_maxLen As Variant, gI_CMP_catename_maxLen As Variant, gI_UDRE_catename_maxLen As Variant, gI_UDRP_catename_maxLen As Variant, gI_UDRP0_catename_maxLen As Variant
Dim gI_UDRP1_catename_maxLen As Variant, gI_CMPE_catename_maxLen As Variant, gI_CMPP_catename_maxLen As Variant, gI_CMPP0_catename_maxLen As Variant, gI_CMPP1_catename_maxLen As Variant
Dim gL_eFuse_catename_maxLen As Variant

    FuseType = UCase(Trim(FuseType))

    If (FuseType = "ECID") Then
        m_len = gI_ECID_catename_maxLen

    ElseIf (FuseType = "CFG") Then
        m_len = gI_CFG_catename_maxLen

    ElseIf (FuseType = "UID") Then
        m_len = gI_UID_catename_maxLen

    ElseIf (FuseType = "UDR") Then
        m_len = gI_UDR_catename_maxLen

    ElseIf (FuseType = "SEN") Then
        m_len = gI_SEN_catename_maxLen
        
    ElseIf (FuseType = "MON") Then
        m_len = gI_MON_catename_maxLen

    ElseIf (FuseType = "CMP") Then
        m_len = gI_CMP_catename_maxLen

    ElseIf (FuseType = "UDRE") Then
        m_len = gI_UDRE_catename_maxLen

    ElseIf (FuseType = "UDRP") Then
        m_len = gI_UDRP_catename_maxLen
        
    ElseIf (FuseType = "UDRP0") Then
        m_len = gI_UDRP0_catename_maxLen
        
    ElseIf (FuseType = "UDRP1") Then
        m_len = gI_UDRP1_catename_maxLen
    ElseIf (FuseType = "CMPE") Then
        m_len = gI_CMPE_catename_maxLen

    ElseIf (FuseType = "CMPP") Then
        m_len = gI_CMPP_catename_maxLen
        
    ElseIf (FuseType = "CMPP0") Then
        m_len = gI_CMPP0_catename_maxLen
        
    ElseIf (FuseType = "CMPP1") Then
        m_len = gI_CMPP1_catename_maxLen

    Else
        theexec.Datalog.WriteComment "auto_eFuse_GetCatenameMaxLen:: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If

    If (gL_eFuse_catename_maxLen = 0) Then
        auto_eFuse_GetCatenameMaxLen = m_len
    Else
        auto_eFuse_GetCatenameMaxLen = gL_eFuse_catename_maxLen  ''''20150702, to align the datalog once mixed fuse, ex, TMPS
    End If

    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + FormatNumeric("Fuse GetCatenameMaxLen", -25)
        m_dlogstr = vbTab & FuseType + "MaxLength = " + FormatNumeric(m_len, -5)
        theexec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_eFuse_GetCatenameMaxLen")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210629, Modify for pattern should expand dssc src wave
Public Function eFuse_DSSC_SetupDigSrcWave_allSites(patt As String, DigSrcPin As PinList, SignalName As String, srcWave As DSPWave, Optional enableExpand As Boolean = False, Optional Expandbit As Long = 0)
On Error GoTo errHandler
Dim InWave As New DSPWave
Dim waveDblArray() As Double
Dim site As Variant
Dim WaveDef As String
Dim m_segsize As Long
'20210629, Modify for pattern should expand dssc src wave
Dim srcWaveExpand As New DSPWave
Dim m_segsizeExpand As Long
Dim i As Long, j As Long

    For Each site In theexec.sites
        If enableExpand = True Then
            m_segsizeExpand = srcWave.sampleSize * Expandbit
            srcWaveExpand.CreateConstant 0, m_segsizeExpand, srcWave.DataType
            
            For i = 0 To srcWave.sampleSize - 1
                For j = 0 To Expandbit - 1
                    srcWaveExpand.ElementLite(i * Expandbit + j) = srcWave.ElementLite(i)
                Next j
            Next i
            InWave = srcWaveExpand.ConvertDataTypeTo(DspDouble).Copy
        Else
            InWave = srcWave.ConvertDataTypeTo(DspDouble).Copy
        End If
        
        waveDblArray = InWave.data
        m_segsize = InWave.sampleSize
        Exit For
    Next site
    
    WaveDef = "WaveDef_" + SignalName + "_allSites"
    TheHdw.Patterns(patt).Load
    
    ''''<NOTICE> Here waveDblArray() must be Double for this case
    theexec.WaveDefinitions.CreateWaveDefinition WaveDef, waveDblArray, True
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName

    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef
        .sampleSize = m_segsize
        .Amplitude = 1
        ''.LoadSamples ''''could waste TT and break PTE
        If glb_TesterType = "Jaguar" Then .LoadSettings
    End With

    ''''<NOTICE> check if there is already one outside
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "eFuse_DSSC_SetupDigSrcWave_allSites")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210629, Modify for pattern should expand dssc src wave
Public Function eFuse_DSSC_SetupDigSrcWave(patt As String, DigSrcPin As PinList, SignalName As String, ByVal srcWave As DSPWave, Optional enableExpand As Boolean = False, Optional Expandbit As Long = 0)
On Error GoTo errHandler
Dim site As Variant
Dim m_segsize As Long
Dim WaveDef As String
'20210629, Modify for pattern should expand dssc src wave
Dim srcWaveExpand As New DSPWave
Dim m_segsizeExpand As Long
Dim i As Long, j As Long

    TheHdw.Patterns(patt).Load
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName

    For Each site In theexec.sites
        If enableExpand = True Then
            m_segsizeExpand = srcWave.sampleSize * Expandbit
            srcWaveExpand.CreateConstant 0, m_segsizeExpand, srcWave.DataType
            
            For i = 0 To srcWave.sampleSize - 1
                For j = 0 To Expandbit - 1
                    srcWaveExpand.ElementLite(i * Expandbit + j) = srcWave.ElementLite(i)
                Next j
            Next i
            
            m_segsize = m_segsizeExpand
            WaveDef = "WaveDef_" + SignalName + "_" & site
            theexec.WaveDefinitions.CreateWaveDefinition WaveDef, srcWaveExpand, True
        Else
            m_segsize = srcWave.sampleSize
            ''''20170920 <NOTICE> if multiple apply this function call/sequence to avoid the following SrcWave to overwrite the previous one
            WaveDef = "WaveDef_" + SignalName + "_" & site
            theexec.WaveDefinitions.CreateWaveDefinition WaveDef, srcWave, True
        End If

        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
            .WaveDefinitionName = WaveDef
            .sampleSize = m_segsize
            .Amplitude = 1
            ''.LoadSamples ''''could waste TT and break PTE
            If glb_TesterType = "Jaguar" Then .LoadSettings
        End With
    Next site
    TheHdw.Wait 0.0001
    
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
    TheHdw.Wait 0.0001

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "eFuse_DSSC_SetupDigSrcWave")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_Hex2BinStr(ByVal HexStr As String, Optional bitwidth As Long = 0) As String
On Error GoTo errHandler
Dim i As Long
Dim PerChar As String
Dim binstr As String
Dim j As Long
Dim DecodeBin As String
Dim MyArray() As Variant
Dim myArrayBin() As Variant
Dim m_len As Long

    HexStr = UCase(HexStr)
    If (InStr(1, HexStr, "X") = 1) Then
        HexStr = Replace(HexStr, "X", "", 1, 1)
    ElseIf (InStr(1, HexStr, "0X") = 1) Then
        HexStr = Replace(HexStr, "0X", "", 1, 1)
    End If

    MyArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                    "A", "B", "C", "D", "E", "F")

    myArrayBin = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", "1000", "1001", _
                       "1010", "1011", "1100", "1101", "1110", "1111")

    binstr = vbNullString
    For i = 1 To Len(HexStr)
        PerChar = mid(HexStr, i, 1)
        'One-to-One mapping, myarray() mappping to myarraybin()
        For j = 0 To UBound(MyArray)
            If (PerChar = MyArray(j)) Then
               DecodeBin = myArrayBin(j)
               Exit For
            End If
        Next j
        binstr = binstr + DecodeBin
    Next i

    m_len = Len(binstr)
    If (bitwidth <> 0) Then
        If (m_len > bitwidth) Then
            For i = 1 To (m_len - bitwidth)
                PerChar = mid(binstr, 1, 1)
                If (PerChar = "0") Then
                    binstr = Replace(binstr, "0", "", 1, 1)
                Else
                    theexec.AddOutput "<WARNING> " + "auto_Hex2BinStr" + ":: Effect BinStr(" + binstr + ") Length(" & Len(binstr) & ") > BitWidth(" & bitwidth & ")"
                    Call Print_Error_Message(Warning_Info, "LIB_eFuse_Func", "auto_Hex2BinStr", ":: Effect BinStr(" + binstr + ") Length(" & Len(binstr) & ") > BitWidth(" & bitwidth & ")")
                End If
            Next i
        ElseIf (m_len < bitwidth) Then
            For i = 1 To (bitwidth - m_len)
                binstr = "0" + binstr
            Next i
        End If
    End If
    auto_Hex2BinStr = binstr

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_Hex2BinStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_eFuse_PatSetToPat_Validation(ByVal patset As Pattern, ByRef patt As String, Optional Validating_ As Boolean, _
Optional showTime As Boolean = False) As Boolean
On Error GoTo errHandler
Dim funcName As String: funcName = "auto_eFuse_PatSetToPat_Validation"
Dim i As Long
Dim PatAry() As String, patCnt As Long
Dim m_patset As New Pattern
Dim m_pat As String

    If (showTime) Then Call auto_StartWatchTimer
    ''''20171211 add, to prevent Unexpected pattern DSSC with Empty data
    TheHdw.Digital.Patgen.Halt ''''<MUST and Could be>

    ''''------------------------------------------------------------------------
    ''''Parsing PatternSet to get the raw pattern name (.pat)
    ''''Actually, eFuse PatternSet only contains one pat file.
    ''''------------------------------------------------------------------------
    'PatAry = TheExec.DataManager.Raw.GetPatternsInSet(PatSet, patCnt)
    If (LCase(patset.value) Like "*.pat") Then
        ReDim PatAry(0)
        PatAry(0) = patset.value
    Else
        PatAry = theexec.DataManager.Raw.GetPatternsInSet(patset, patCnt)
    End If
    
    While Not (LCase(PatAry(0)) Like "*.pat*")
        m_patset.value = PatAry(0)
        PatAry = theexec.DataManager.Raw.GetPatternsInSet(m_patset, patCnt)
        If UBound(PatAry) > 1 Then theexec.ErrorLogMessage (patset & " is with more than one pattern in the pattern set")
    Wend
    patt = PatAry(0)
    ''''<NOTICE> lowcase ".gz" will be implicit, so pattern(.gz).load/test will be problem.
    ''''20161124 update to prevent *.gz (lowcase)
    If (patt Like "*.gz") Then
        patt = Replace(patt, ".gz", "")
    End If
    ''''------------------------------------------------------------------------

    auto_eFuse_PatSetToPat_Validation = False ''''init
    If (Validating_) Then
        ''''<MUST> By this way, the PatSet can be explicit in the pattern memory.
        If (ValidatePattern(patset.value) = False) Then
            theexec.AddOutput "<Error> " + funcName + ":: please check the PatternSet, " + CStr(patset.value)
            Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_eFuse_PatSetToPat_Validation", funcName + ":: please check the PatternSet, " + CStr(patset.value))
            GoTo errHandler
        End If

        ''''Actually, here it's just only one pattern.
        ''''<MUST> By this way, the individual pattern (.pat) can be explicit in the pattern memory.
        For i = 0 To UBound(PatAry)
            m_pat = PatAry(i)
            If (ValidatePattern(m_pat) = False) Then
                theexec.AddOutput "<Error> " + funcName + ":: please check the Pattern, " + CStr(m_pat)
                Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_eFuse_PatSetToPat_Validation", funcName + ":: please check the Pattern, " + CStr(m_pat))
                GoTo errHandler
            End If
        Next i
        auto_eFuse_PatSetToPat_Validation = True ''''<MUST>
        If (showTime) Then Call auto_StopWatchTimer(funcName)
        Exit Function ''''<MUST>
    End If

    For i = 0 To UBound(PatAry)
        m_pat = PatAry(i)
        ''''<NOTICE> lowcase ".gz" will be implicit, so pattern(.gz).load/test will be problem.
        ''''20161124 update to prevent *.gz (lowcase)
        If (m_pat Like "*.gz") Then
            m_pat = Replace(m_pat, ".gz", "")
        End If
        
        '''<MUST> put it here when user unloadAllPatterns, it can reload the pattern.
        '''<MUST> DSSC Src/Cap setup needs this statement
        TheHdw.Patterns(m_pat).Load ''''it will take 0.3-0.5 ms if the pattern was loaded
    Next i

    TheHdw.Wait 0.00001 ''''10uS
    If (showTime) Then Call auto_StopWatchTimer(funcName)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_eFuse_PatSetToPat_Validation")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TurnOnEfusePwrPins(powerPin As String, _
                                   Optional v As Double = 1.8, _
                                   Optional i_rng As Double = 0.2, _
                                   Optional wait_before_gate As Double = 0.001, _
                                   Optional wait_after_gate As Double = 0.002, _
                                   Optional Steps As Long = 10, _
                                   Optional RiseTime As Double = 0.001)

On Error GoTo errHandler
Dim m_currVolt As Double

    '******************** 2015/1/16 Laba ********************
    'ECID Supply : VDD18_EFUSE0
    'HDCP Keys, UID Keys, Sensor Trim Values, Config : VDD18_EFUSE1
    'SOC BIRA1, SOC BIRA2, CPU UDR+BIRA, GFX BIRA : VDD18_EFUSE2

    ''''auto_eFuse_pwr_on_i_meter_DCVS(pin As String, v As Double, i_rng As Double, wait_before_gate As Double, wait_after_gate As Double, steps As Long, RiseTime As Double)
    ''''auto_eFuse_pwr_on_i_meter_DCVS PowerPin, vpwr, 0.2, 0.001, 0.002, 10, 0.001
    auto_eFuse_pwr_on_i_meter_DCVS powerPin, v, i_rng, wait_before_gate, wait_after_gate, Steps, RiseTime

    m_currVolt = TheHdw.DCVS.Pins(powerPin).Voltage.Main.value
    theexec.Datalog.WriteComment ""
    theexec.Datalog.WriteComment "TurnOnEfusePwrPins :: " + UCase(powerPin) + " = " + Format(m_currVolt, "0.000")
    theexec.Datalog.WriteComment ""

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "TurnOnEfusePwrPins")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TurnOffEfusePwrPins(powerPin As String, _
                                   Optional v As Double = 1.8, _
                                   Optional i_rng As Double = 0.2, _
                                   Optional wait_before_gate As Double = 0.001, _
                                   Optional wait_after_gate As Double = 0.002, _
                                   Optional Steps As Long = 10, _
                                   Optional RiseTime As Double = 0.001)

On Error GoTo errHandler
Dim m_currVolt As Double
    ''auto_eFuse_pwr_off_i_meter_DCVS CurrentVoltage, 1.8, 0.2, 0.001, 0.002, 10, 0.001

    auto_eFuse_pwr_off_i_meter_DCVS powerPin, v, i_rng, wait_before_gate, wait_after_gate, Steps, RiseTime

    m_currVolt = TheHdw.DCVS.Pins(powerPin).Voltage.Main.value
    theexec.Datalog.WriteComment ""
    theexec.Datalog.WriteComment "TurnOffEfusePwrPins :: " + UCase(powerPin) + " = " + Format(m_currVolt, "0.000")
    theexec.Datalog.WriteComment ""

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "TurnOffEfusePwrPins")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_StartWatchTimer()
On Error GoTo errHandler

    TheHdw.StartStopwatch

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_StartWatchTimer")
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20161114 update
Public Function auto_StopWatchTimer(Optional itemStr As String = vbNullString, Optional unit_ms As Boolean = True)
On Error GoTo errHandler
Dim m_instName As String
Dim m_tmpStr As String
Dim m_timedbl As Double
Dim m_len As Long

    m_timedbl = TheHdw.ReadStopwatch
    
    m_instName = FormatNumeric(theexec.DataManager.instancename, 35)
    itemStr = FormatNumeric(itemStr, 10)

    If (Trim(m_instName) <> "" And Trim(itemStr) <> "") Then
        m_tmpStr = m_instName + "::" + itemStr + " = "
    ElseIf (Trim(m_instName) = "" And Trim(itemStr) <> "") Then
        m_tmpStr = itemStr + " = "
    ElseIf (Trim(m_instName) <> "" And Trim(itemStr) = "") Then
        m_tmpStr = m_instName + " = "
    Else
        m_tmpStr = vbNullString
    End If
    m_len = Len(m_tmpStr)
    
    If (unit_ms) Then
        m_tmpStr = vbTab & "Test Time " + FormatNumeric(m_tmpStr, m_len) + Format(m_timedbl * 1000, "0.0000") + " mS."
    Else
        m_tmpStr = vbTab & "Test Time " + FormatNumeric(m_tmpStr, m_len) + Format(m_timedbl, "0.000000") + " Secs."
    End If
    ''TheExec.Datalog.WriteComment m_tmpStr
    Debug.Print m_tmpStr
    ''''If (UCase(m_tmpStr) Like UCase("*Validation*")) Then Debug.Print m_tmpStr
    
    ''''using for the next/continue stopWatch to avoid the above code statements (extral time)
    TheHdw.StartStopwatch
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_StopWatchTimer")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_eFuse_pwr_on_i_meter_DCVS(pin As String, v As Double, i_rng As Double, _
                                               wait_before_gate As Double, wait_after_gate As Double, _
                                               Steps As Long, RiseTime As Double, _
                                               Optional showPrint As Boolean = False)

On Error GoTo errHandler
Dim i_meter_rng As Double
Dim setV As Double
Dim StepV As Double
Dim i As Long
Dim stepT As Double
    'set Force voltage and Current/Meter Range
    ''===============================================
    ''Description: __                __
    ''                __|
    ''            __|
    ''        __|
    ''    __| >|   |<--stepT  __        v
    ''__|                   __ stepV   __
    ''|<-- steps -->
    ''|<-FallTime ->
    ''===============================================

    i_meter_rng = i_rng
    StepV = v / Steps
    stepT = RiseTime / Steps

    With TheHdw.DCVS.Pins(pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = 0
        .SetCurrentRanges i_rng, i_meter_rng
'        .CurrentLimit.Source.FoldLimit.Level = i_rng
        .Meter.mode = tlDCVSMeterCurrent
        .CurrentRange.value = i_rng
        .CurrentLimit.Source.FoldLimit.level.value = i_rng
        If glb_TesterType = "Jaguar" Then .Meter.CurrentRange = i_rng
        TheHdw.Wait wait_before_gate   'wait for relay connect
        .Gate = True
    End With
    
    ''Pwr On Ramp up slew-rate control============================
    For i = 1 To Steps
        setV = i * StepV
        TheHdw.DCVS.Pins(pin).Voltage.Main = setV
        
        If showPrint = True Then
            theexec.Datalog.WriteComment "  Curr_" & pin & " Pwr Up Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        
        TheHdw.Wait stepT
    Next i
    ''============================================================

    TheHdw.Wait wait_after_gate

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_eFuse_pwr_on_i_meter_DCVS")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_eFuse_pwr_off_i_meter_DCVS(pin As String, v As Double, i_rng As Double, _
                                                wait_before_gate As Double, wait_after_gate As Double, _
                                                Steps As Long, FallTime As Double, _
                                                Optional showPrint As Boolean = False)

On Error GoTo errHandler
Dim i_meter_rng As Double
Dim setV As Double
Dim StepV As Double
Dim i As Long
Dim stepsm As Long
Dim stepT As Double
    'set Force voltage and Current/Meter Range
    ''===============================================
    ''Description
    ''__                              __
    ''  |__
    ''     |_>|  |<--stepT             v
    ''        |__          __
    ''           |__       __ stepV   __
    ''|<-- steps -->
    ''|<-FallTime ->
    ''===============================================

    i_meter_rng = i_rng
    StepV = v / Steps
    stepT = FallTime / Steps

    With TheHdw.DCVS.Pins(pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = v
        .SetCurrentRanges i_rng, i_meter_rng
'        .CurrentLimit.Source.FoldLimit.Level = i_rng
        .Meter.mode = tlDCVSMeterCurrent
        .CurrentRange.value = i_rng
        .CurrentLimit.Source.FoldLimit.level.value = i_rng
        If glb_TesterType = "Jaguar" Then .Meter.CurrentRange = i_rng
        TheHdw.Wait wait_before_gate   'wait for relay connect
        .Gate = True
    End With
    
    ''Pwr On Ramp Down slew-rate control============================
    stepsm = Steps - 1
    For i = 0 To stepsm
        setV = v - (i * StepV)
        TheHdw.DCVS.Pins(pin).Voltage.Main = setV
        
        If showPrint = True Then
            theexec.Datalog.WriteComment "  Curr_" & pin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        
        TheHdw.Wait stepT
    Next i

    setV = 0
    TheHdw.DCVS.Pins(pin).Voltage.Main = setV
    
    If showPrint = True Then
        theexec.Datalog.WriteComment "  Curr_" & pin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
    End If
    ''==============================================================

    TheHdw.Wait wait_after_gate

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "auto_eFuse_pwr_off_i_meter_DCVS")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function eFuse_DSSC_SetupDigSrcArr_allSites(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, WaveDefArray() As Long)
On Error GoTo errHandler
Dim InWave As New DSPWave
Dim waveDblArray() As Double
Dim site As Variant
Dim WaveDef As String

    For Each site In theexec.sites.Active
        InWave.data = WaveDefArray
        InWave = InWave.ConvertDataTypeTo(DspDouble)
        waveDblArray = InWave.data
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
        If glb_TesterType = "Jaguar" Then .LoadSettings
    End With

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "eFuse_DSSC_SetupDigSrcArr_allSites")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DSSC_SetupDigSrcWave_TTR(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, InWave As DSPWave)
On Error GoTo errHandler
Dim site As Variant
Dim WaveDef As String
    'store efuse program bit into a DSP wave
    WaveDef = "WaveDef"
    site = theexec.sites.siteNumber

    TheHdw.Patterns(patt).Load
    theexec.WaveDefinitions.CreateWaveDefinition WaveDef & site, InWave, True
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef & site
        .sampleSize = SegmentSize
        .Amplitude = 1
        .LoadSamples
        If glb_TesterType = "Jaguar" Then .LoadSettings
    End With
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "DSSC_SetupDigSrcWave_TTR")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190507: Added "Cdec" to avoid double format accuracy issues
'20181207: Added "EnableFuse" to decide fusing or not...
'20181011: Added for postbincut
'20180919: Modified for the new parsing method
'20180830: Added for SEPVM
Public Function SEPVM_Fuse_mod(p_mode As Integer, ByRef pmode_val As SiteDouble, Sepvm_table As SEPVM_SHEET_Type, Optional EnableFuse As Boolean = False)
On Error GoTo errHandler
Dim Row As Long
Dim Col As Long
Dim start_row As Long
Dim Start_col As Long
Dim split_content() As String
Dim eqn_count As Long
Dim site As Variant
Dim limit_cnt As Long
Dim range_found As New SiteBoolean
Dim FuseVal As New SiteDouble
Dim opbank As eFuseBdfBank ''20201112 add for obj
Dim field As eFuseBdfField
    'Dim inst_name As String 'added for postbincut, 20181011
    'Dim ws_def As Worksheet
    'Dim wb As Workbook

    'Set wb = Application.ActiveWorkbook
    'Set ws_def = wb.Sheets("SEPVM_fusing")

    eqn_count = 0
    'inst_name = LCase(Trim(TheExec.DataManager.instanceName))

    '''Find the range
    For Each site In theexec.sites
        '''init
        range_found(site) = False
        FuseVal(site) = 0
        
        '''//Note:
        '''//pmode_val is the product voltage of p_mode.
        '''//Look up the sheet "SEPVM_fusing" and decide voltage range, then get SEPVTH.
        '''//voltage=max((pmode_val*1.1), (SEPVTH+18mv))
        '''Added "Cdec" to avoid double format accuracy issues, 20190507
        If CDec(pmode_val(site)) > 0 Then
            For limit_cnt = 0 To Sepvm_table.RangeCnt 'modified for SEPVM,  cnt:17 means 0~16, 20181005
                If (pmode_val(site) >= Sepvm_table.Sepvm_Const(limit_cnt).Lo_Limit) And (pmode_val(site) < Sepvm_table.Sepvm_Const(limit_cnt).Hi_Limit) Then 'Modified for the parsing method and data type, 20180919
''                    pmode_val = IIf(((pmode_val * 1.1) > (Sepvm_table.Sepvm_Const(limit_cnt).Sepvm_vth + 18)), (pmode_val * 1.1), (Sepvm_table.Sepvm_Const(limit_cnt).Sepvm_vth + 18))
''                    cal_corepower_val = pmode_val
                    range_found(site) = True
                    theexec.Datalog.WriteComment "site(" & site & ")" & VddBinName(p_mode) & " = " & pmode_val(site) & " mV"
                    theexec.Datalog.WriteComment "site(" & site & ")" & Sepvm_table.FuseName & " = " & Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0
'                    TheExec.Datalog.WriteComment Sepvm_table.FuseName(0) & " = " & Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0
'                    TheExec.Datalog.WriteComment Sepvm_table.FuseName(1) & " = " & Sepvm_table.Sepvm_Const(limit_cnt).FuseValue1
                    
                    '''//**************************************************************//'''
                    '''Added "EnableFuse" to decide fusing for SEPVM, 20181207
                    '''Modified for SEPVM, 20181011. In postBinCut, we don't do sepvm...
                    '''"EnableFuse" is decided in HVCC_Set_VT or PostBinCut_Voltage_Set_VT
                    '''<old>
                        'If inst_name Like "*_hbv*" And Not (inst_name Like "*_binresult_*") Then
                    '''<new>
                    If EnableFuse Then
                        If LCase(Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0) Like "not valid" Then
                        
                         theexec.sites.item(site).FlagState("F_Efuse_PreWrite") = logicTrue 'bin out "not valid" case from Jose
                        
                        Else
                            FuseVal(site) = Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0
'                            Call auto_eFuse_SetWriteDecimal("MON", Sepvm_table.FuseName, Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0, True)
                        End If
                    End If
                    '''//**************************************************************//'''
                    Exit For
                End If
            Next limit_cnt
            
            If (Not range_found(site)) Or LCase(Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0) Like "not valid" Then
                'theexec.DataLog.WriteComment "No proper range for SEPVM voltage calculation!!!"
                theexec.ErrorLogMessage VddBinName(p_mode) & "," & site & ", Fuse value is no valid:(" & pmode_val(site) & ") No proper voltage range in SRAM_SOC_fusing table!!!"
                'GoTo errHandler
                theexec.Flow.TestLimit resultVal:=999, lowVal:=0, hiVal:=0, Tname:="PreFuse," & site 'modified for BinOut, 20180918
                
                theexec.sites.item(site).FlagState("F_Efuse_PreWrite") = logicTrue 'bin out "not valid" case from Jose
                
            Else
                 theexec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:="PreFuse," & site
            End If
        End If
    Next site
    
    'Call auto_eFuse_SetWriteVariable_SiteAware("MON", Sepvm_table.FuseName, FuseVal, True)
    ''====20201230 add for efuse new code====
    Set opbank = GetBdfBank("MON")
    Set field = opbank.Fields(Sepvm_table.FuseName)
    opbank.SetEfuse field.name, FuseVal, , , , , True
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "SEPVM_Fuse_mod")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190507: Added "Cdec" to avoid double format accuracy issues
'20181207: Added "EnableFuse" to decide fusing or not...
'20181011: Added for postbincut
'20180919: Modified for the new parsing method
'20180830: Added for SEPVM
'20220614, Add for new sram_soc_fusingXXX parsing
Public Function Parsing_SEPVM_Fusing_Table_dynamic(tablekeyword As String) As Boolean
On Error GoTo errHandler
Dim wb As Workbook
Dim ws_def As Worksheet
Dim i As Long
Dim MaxRow As Long, maxcol As Long, RowRange As Long, ColRange As Long, startRow, startCol As Long, eqn_count As Long
Dim split_content() As String
Dim tmp_str As String, keyword As String
Dim product_range_col As Long, FuseValue0_col As Long, RangeFound As Long
Dim Sepvm_table As SEPVM_SHEET_Type
Dim strAry_allsheetnames As Variant
Dim strAry_usingsheetnames() As String
Dim idxSheet As Integer, usingsheetcnt As Integer, iRow As Integer
Dim LowestCheck_Val As Double, HighestCheck_Val As Double, SepvmCheck_Val As Double
Dim CheckPass As Boolean
Dim stempValue As String: stempValue = vbNullString
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim fuseValue As New SiteVariant
Dim valueStr As String: valueStr = vbNullString
Dim msg As String

    Parsing_SEPVM_Fusing_Table_dynamic = True
    '''//Please fined the definition of sheetType in "DataTool.DMGRSheetTypes".
    strAry_allsheetnames = theexec.job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)

    usingsheetcnt = 0
    For idxSheet = 0 To UBound(strAry_allsheetnames)
        If LCase(strAry_allsheetnames(idxSheet)) Like LCase(tablekeyword & "*") Then
            ReDim Preserve strAry_usingsheetnames(0 To usingsheetcnt)
            strAry_usingsheetnames(usingsheetcnt) = strAry_allsheetnames(idxSheet)
            usingsheetcnt = usingsheetcnt + 1
        End If
    Next idxSheet

    If usingsheetcnt > 0 Then
        Set opbank = GetBdfBank("mon")
        ReDim Sepvm_table_Arr(UBound(strAry_usingsheetnames))
        For idxSheet = 0 To UBound(strAry_usingsheetnames)
            '''Init
            msg = vbNullString
            eqn_count = 0
            product_range_col = -1
            FuseValue0_col = -1
            
            '''start point
            '''    row = 1
            '''    col = 1
            Set wb = Application.ActiveWorkbook
            Set ws_def = wb.Sheets(strAry_usingsheetnames(idxSheet))
            
            ''''''''''''''''''''''''''''''''''''''''''''''
            '''//Find the start cell//
            '''Modified for parsing sheet "SEPVM_fusing", 20180907
            MaxRow = ws_def.Cells.SpecialCells(xlCellTypeLastCell).Row
            maxcol = ws_def.Cells.SpecialCells(xlCellTypeLastCell).Column
            RowRange = ws_def.UsedRange.Rows.Count
            ColRange = ws_def.UsedRange.Columns.Count
            startRow = MaxRow - RowRange + 1
            startCol = maxcol - ColRange + 1
            
            '''Initialize structure "Sepvm_table"
            Sepvm_table.FuseName = vbNullString
            Sepvm_table.RangeCnt = 0
            Erase Sepvm_table.Sepvm_Const()
            
            '''Search the p_mode, find the columns for headers. Modified on 20180919
            tmp_str = vbNullString
            For i = 0 To ColRange - 1
                If LCase(Trim(ws_def.Cells(startRow, startCol + i).value)) Like "*value*" Then
                    product_range_col = startCol + i
                ElseIf LCase(Trim(ws_def.Cells(startRow, startCol + i).value)) Like "*mon*" Then
                    FuseValue0_col = startCol + i
                    tmp_str = LCase(Trim(ws_def.Cells(startRow, startCol + i).value))
                End If
            Next i
            Sepvm_table.FuseName = tmp_str
            If opbank.Fields.Exists(Sepvm_table.FuseName) Then
                Set field = opbank.Fields(Sepvm_table.FuseName)
            Else
                msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet fuse category name not exist in BDF table!!!"
                GoTo SkipAndPrintErrorMsg
            End If
            If product_range_col > 0 And FuseValue0_col > 0 Then
                startRow = startRow + 1
            Else
                'TheExec.AddOutput "The SEPVM_fusing sheet doesn't have correct headers!!! Error!!!"
                msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet doesn't have correct headers!!! Error!!!"
                GoTo SkipAndPrintErrorMsg
            End If
            
            '''Start parsing the whole sheet
            '''Find the range
            For iRow = startRow To MaxRow
                valueStr = LCase(ws_def.Cells(iRow, product_range_col).value)
                If valueStr Like "*>*" Then
                    msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet not allow a math symbol >, please check it (row:" + CStr(iRow) + " column:" + CStr(product_range_col) + ")"
                    GoTo SkipAndPrintErrorMsg
                ElseIf valueStr Like "*<*" Then 'Modified for the latest SEPVM_fusing, 20180928
                    '''Each range should show "lo_limit <= product <= hi_limit"
                    ReDim Preserve Sepvm_table.Sepvm_Const(0 To eqn_count)
                    split_content = Split(LCase(Trim(ws_def.Cells(iRow, product_range_col).value)), "<")
                    
                    '----------------Find the Keyword-----------------------------------------------
                    tmp_str = vbNullString
                    For i = 0 To UBound(split_content)
                        tmp_str = Trim(Replace(split_content(i), "=", ""))
                        If IsNumeric(tmp_str) = False Then keyword = tmp_str
                    Next i
                    '----------------Find the Keyword-----------------------------------------------
                    split_content = Split(LCase(Trim(ws_def.Cells(iRow, product_range_col).value)), keyword)
                    If split_content(0) = "" Then split_content(0) = "0"
                    If split_content(1) = "" Then split_content(1) = "9999"
                    stempValue = vbNullString
                    ''''''parsing Sepvm th value
                    ''' handle Low limit
                    If InStr(split_content(0), "<") <> 0 Then
                        stempValue = Replace(split_content(0), "<", "")
                        If InStr(stempValue, "=") <> 0 Then
                            stempValue = Trim(Replace(stempValue, "=", ""))
                            Sepvm_table.Sepvm_Const(eqn_count).Lo_Limit = CDbl(stempValue)
                        Else
                            Sepvm_table.Sepvm_Const(eqn_count).Lo_Limit = CDbl(stempValue) + 0.000001 'no =
                        End If
                    Else
                        If InStr(split_content(0), "=") <> 0 Then
                            stempValue = Trim(Replace(split_content(0), "=", ""))
                            Sepvm_table.Sepvm_Const(eqn_count).Lo_Limit = CDbl(stempValue)
                        Else
                            Sepvm_table.Sepvm_Const(eqn_count).Lo_Limit = split_content(0)
                        End If
                    End If
                    ''' handle High limit
                    If InStr(split_content(1), "<") <> 0 Then
                        stempValue = Replace(split_content(1), "<", "")
                        If InStr(stempValue, "=") <> 0 Then
                            stempValue = Trim(Replace(stempValue, "=", ""))
                            Sepvm_table.Sepvm_Const(eqn_count).Hi_Limit = CDbl(stempValue)
                        Else
                            Sepvm_table.Sepvm_Const(eqn_count).Hi_Limit = CDbl(stempValue) - 0.000001 'no =
                        End If
                    Else
                        If InStr(split_content(1), "=") <> 0 Then
                            stempValue = Trim(Replace(split_content(1), "=", ""))
                            Sepvm_table.Sepvm_Const(eqn_count).Hi_Limit = CDbl(stempValue)
                        Else
                            Sepvm_table.Sepvm_Const(eqn_count).Hi_Limit = split_content(1)
                        End If
                    End If
                    
                    '''parsing Sepvm fuse value
                    If Trim(ws_def.Cells(iRow, FuseValue0_col).value) = "" Then
                       msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet fuse value is empty, please check it (row:" + CStr(iRow) + " column:" + CStr(FuseValue0_col) + ")"
                       GoTo SkipAndPrintErrorMsg
                    Else
                        If LCase(Trim(ws_def.Cells(iRow, FuseValue0_col).value)) Like "not valid" Then
                            Sepvm_table.Sepvm_Const(eqn_count).FuseValue0 = LCase(Trim(ws_def.Cells(iRow, FuseValue0_col).value))
                        Else
                            Sepvm_table.Sepvm_Const(eqn_count).FuseValue0 = CDbl(LCase(Trim(ws_def.Cells(iRow, FuseValue0_col).value)))
                            fuseValue = Sepvm_table.Sepvm_Const(eqn_count).FuseValue0
                            fuseValue = GlbUtility.oDec2HexStr(fuseValue, field.hBytes)
                            If field.IsOverflow(fuseValue, bMsgBox:=False) Then
                                msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet fuse value is overflow, please check it (row:" + CStr(iRow) + " column:" + CStr(FuseValue0_col) + ")"
                                GoTo SkipAndPrintErrorMsg
                            End If
                        End If
                    End If

                    Sepvm_table.RangeCnt = eqn_count
                    eqn_count = eqn_count + 1
                ElseIf valueStr Like "*=*" Then
                    ReDim Preserve Sepvm_table.Sepvm_Const(0 To eqn_count)
                    split_content = Split(LCase(Trim(ws_def.Cells(iRow, product_range_col).value)), "=")
                    '----------------Find the Keyword-----------------------------------------------
                    tmp_str = vbNullString
                    For i = 0 To UBound(split_content)
                        tmp_str = Trim(Replace(split_content(i), "=", ""))
                        If IsNumeric(tmp_str) = True Then keyword = tmp_str
                    Next i
                    '----------------Find the Keyword-----------------------------------------------
                    Sepvm_table.Sepvm_Const(eqn_count).Lo_Limit = keyword
                    Sepvm_table.Sepvm_Const(eqn_count).Hi_Limit = keyword
                    
                    If Trim(ws_def.Cells(iRow, FuseValue0_col).value) = "" Then
                       msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet fuse value is empty, please check it (row:" + CStr(iRow) + " column:" + CStr(FuseValue0_col) + ")"
                       GoTo SkipAndPrintErrorMsg
                    Else
                        If LCase(Trim(ws_def.Cells(iRow, FuseValue0_col).value)) Like "not valid" Then
                            Sepvm_table.Sepvm_Const(eqn_count).FuseValue0 = LCase(Trim(ws_def.Cells(iRow, FuseValue0_col).value))
                        Else
                            Sepvm_table.Sepvm_Const(eqn_count).FuseValue0 = CDbl(LCase(Trim(ws_def.Cells(iRow, FuseValue0_col).value)))
                            fuseValue = Sepvm_table.Sepvm_Const(eqn_count).FuseValue0
                            fuseValue = GlbUtility.oDec2HexStr(fuseValue, field.hBytes)
                            If field.IsOverflow(fuseValue, bMsgBox:=False) Then
                                msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet fuse value is overflow, please check it (row:" + CStr(iRow) + " column:" + CStr(FuseValue0_col) + ")"
                                GoTo SkipAndPrintErrorMsg
                            End If
                        End If
                    End If
                    Sepvm_table.RangeCnt = eqn_count
                    eqn_count = eqn_count + 1
                ElseIf Not IsEmpty(valueStr) Then
                    'unexecpted format
                    If valueStr Like "*>*" Then
                        msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet not allow a math symbol >, please check it (row:" + CStr(iRow) + " column:" + CStr(product_range_col) + ")"
                    Else
                        msg = "The " + strAry_usingsheetnames(idxSheet) + " sheet have format issue, please check it (row:" + CStr(iRow) + " column:" + CStr(product_range_col) + ")"
                    End If
                    GoTo SkipAndPrintErrorMsg
                Else
                    'cell is empty
                    Call Print_Error_Message(Warning_Info, "LIB_eFuse_Func", "Parsing_SEPVM_Fusing_Table_dynamic", "The " + strAry_usingsheetnames(idxSheet) + " sheet product value is empty (row:" + CStr(iRow) + " column:" + CStr(product_range_col) + ")")
                End If
            Next iRow
            
            '---------------------------------SanityCheck-----------------------------------
            ''' Lo_limit of current range must be equal to Hi_limit of previous range
            LowestCheck_Val = Replace(Replace(Sepvm_table.Sepvm_Const(0).Hi_Limit, "<", ""), "=", "") - 1
            HighestCheck_Val = Replace(Replace(Sepvm_table.Sepvm_Const(Sepvm_table.RangeCnt).Lo_Limit, "<", ""), "=", "") + 1
            
            If HighestCheck_Val > LowestCheck_Val Then
                SepvmCheck_Val = LowestCheck_Val
                RangeFound = 0
                While SepvmCheck_Val < HighestCheck_Val
                    CheckPass = False
                    For i = RangeFound To Sepvm_table.RangeCnt
                        If (Evaluate(Sepvm_table.Sepvm_Const(i).Lo_Limit - SepvmCheck_Val) >= 0) And (Evaluate(SepvmCheck_Val - Sepvm_table.Sepvm_Const(i).Hi_Limit) < 0) Then
                            If CheckPass = False Then
                                CheckPass = Pass
                                RangeFound = i
                            Else
                                CheckPass = Pass
                            End If
                        ElseIf (Sepvm_table.Sepvm_Const(i).Lo_Limit = Sepvm_table.Sepvm_Const(i).Hi_Limit) Then
                            CheckPass = Pass
                        Else
                            If i <> 0 Then
                                CheckPass = False
                                Exit For
                            End If
                        End If
                        SepvmCheck_Val = Sepvm_table.Sepvm_Const(i).Hi_Limit
                        'If i = RangeFound + 1 Then Exit For
                    Next i
                    If CheckPass = False Then
                        msg = "VDD range of Sheet SEPVM_fusing row(" & CStr(i) & ") is incorrect! Error!!!"
                        GoTo SkipAndPrintErrorMsg
                    Else
                        SepvmCheck_Val = SepvmCheck_Val + 1
                    End If
                Wend
            Else
                msg = "VDD range of Sheet " & strAry_usingsheetnames(idxSheet) & " table format is incorrect! Error!!!"
                GoTo SkipAndPrintErrorMsg
            End If
            
            '---------------------------------SanityCheck-----------------------------------
            If SepvmCheck_Val <> 9999 Then Sepvm_table_Arr(idxSheet) = Sepvm_table
        Next idxSheet
    End If
    
Exit Function
SkipAndPrintErrorMsg:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "Parsing_SEPVM_Fusing_Table_dynamic", msg)
    theexec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="Parsing_SEPVM_Fusing_Table_dynamic"
    Parsing_SEPVM_Fusing_Table_dynamic = False
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "Parsing_SEPVM_Fusing_Table_dynamic")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210624 Carter
Public Function SEPVM_Fuse_mod_AllCaseSupport2(p_mode As Integer, ByRef pmode_val As SiteDouble, Sepvm_table As SEPVM_SHEET_Type, Optional EnableFuse As Boolean = False)
On Error GoTo errHandler
Dim i As Long
Dim Row As Long
Dim Col As Long
Dim start_row As Long
Dim Start_col As Long
Dim split_content() As String
Dim eqn_count As Long
Dim site As Variant
Dim limit_cnt As Long
Dim range_found As New SiteBoolean
Dim FuseVal As New SiteDouble
'Dim inst_name As String 'added for postbincut, 20181011
'Dim ws_def As Worksheet
'Dim wb As Workbook
Dim opbank As eFuseBdfBank ''20201112 add for obj
Dim field As eFuseBdfField

    'Set wb = Application.ActiveWorkbook
    'Set ws_def = wb.Sheets("SEPVM_fusing")
    
    eqn_count = 0
    'inst_name = LCase(Trim(TheExec.DataManager.instanceName))
    
    '''Find the range
    For Each site In theexec.sites
        '''init
        range_found(site) = False
        FuseVal(site) = 0
'        pmode_val(0) = 830.1
'        pmode_val(1) = 850
'        pmode_val(2) = 822.01
        
        '''//Note:
        '''//pmode_val is the product voltage of p_mode.
        '''//Look up the sheet "SEPVM_fusing" and decide voltage range, then get SEPVTH.
        '''//voltage=max((pmode_val*1.1), (SEPVTH+18mv))
        '''Added "Cdec" to avoid double format accuracy issues, 20190507
        If CDec(pmode_val(site)) > 0 Then
            For limit_cnt = 0 To Sepvm_table.RangeCnt 'modified for SEPVM,  cnt:17 means 0~16, 20181005
                If (pmode_val(site) >= Sepvm_table.Sepvm_Const(limit_cnt).Lo_Limit) And (pmode_val(site) <= Sepvm_table.Sepvm_Const(limit_cnt).Hi_Limit) Then
''                    range_found(site) = True
''                    TheExec.Datalog.WriteComment "site(" & site & ")" & VddBinName(p_mode) & " = " & pmode_val(site) & " mV"
''                    TheExec.Datalog.WriteComment "site(" & site & ")" & Sepvm_table.FuseName & " = " & Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0
                    If EnableFuse Then
                        If LCase(Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0) Like "not valid" Then
                            If limit_cnt < Sepvm_table.RangeCnt Then
                                For i = limit_cnt + 1 To Sepvm_table.RangeCnt
                                    If (pmode_val(site) >= Sepvm_table.Sepvm_Const(i).Lo_Limit) And (pmode_val(site) <= Sepvm_table.Sepvm_Const(i).Hi_Limit) Then
                                        range_found(site) = True
                                        theexec.Datalog.WriteComment "site(" & site & ")" & VddBinName(p_mode) & " = " & pmode_val(site) & " mV"
                                        theexec.Datalog.WriteComment "site(" & site & ")" & Sepvm_table.FuseName & " = " & Sepvm_table.Sepvm_Const(i).FuseValue0
                                        If LCase(Sepvm_table.Sepvm_Const(i).FuseValue0) Like "not valid" Then
                                            theexec.sites.item(site).FlagState("F_Efuse_PreWrite") = logicTrue 'bin out "not valid" case from Jose
                                        Else
                                            FuseVal(site) = Sepvm_table.Sepvm_Const(i).FuseValue0
                                        End If
                                        limit_cnt = i
                                        Exit For
                                    End If
                                Next i
                            Else
                                range_found(site) = True
                                theexec.Datalog.WriteComment "site(" & site & ")" & VddBinName(p_mode) & " = " & pmode_val(site) & " mV"
                                theexec.Datalog.WriteComment "site(" & site & ")" & Sepvm_table.FuseName & " = " & Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0
                                theexec.sites.item(site).FlagState("F_Efuse_PreWrite") = logicTrue 'bin out "not valid" case from Jose
                            End If
                        Else
                            range_found(site) = True
                            theexec.Datalog.WriteComment "site(" & site & ")" & VddBinName(p_mode) & " = " & pmode_val(site) & " mV"
                            theexec.Datalog.WriteComment "site(" & site & ")" & Sepvm_table.FuseName & " = " & Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0
                            FuseVal(site) = Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0
                        End If
                    End If
                    Exit For
                End If
            Next limit_cnt
            
            If (Not range_found(site)) Or LCase(Sepvm_table.Sepvm_Const(limit_cnt).FuseValue0) Like "not valid" Then
                'theexec.DataLog.WriteComment "No proper range for SEPVM voltage calculation!!!"
                theexec.ErrorLogMessage VddBinName(p_mode) & "," & site & ", Fuse value is no valid:(" & pmode_val(site) & ") No proper voltage range in SRAM_SOC_fusing table!!!"
                'GoTo errHandler
                theexec.Flow.TestLimit resultVal:=999, lowVal:=0, hiVal:=0, Tname:="PreFuse," & site 'modified for BinOut, 20180918
                
                theexec.sites.item(site).FlagState("F_Efuse_PreWrite") = logicTrue 'bin out "not valid" case from Jose
                
            Else
                 theexec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:="PreFuse," & site
            End If
        End If
    Next site
    
    'Call auto_eFuse_SetWriteVariable_SiteAware("MON", Sepvm_table.FuseName, FuseVal, True)
    ''====20201230 add for efuse new code====
    Set opbank = GetBdfBank("MON")
    Set field = opbank.Fields(Sepvm_table.FuseName)
    If UCase(Replace(BdfDataBase.GetRealStage(field.BlowLocation), "_EARLY", "")) Like GlbUtility.currStage Then 'JackChou 202405
        opbank.SetEfuse field.name, FuseVal, , , , , True
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_eFuse_Func", "SEPVM_Fuse_mod_AllCaseSupport2")
    If AbortTest Then Exit Function Else Resume Next
End Function


