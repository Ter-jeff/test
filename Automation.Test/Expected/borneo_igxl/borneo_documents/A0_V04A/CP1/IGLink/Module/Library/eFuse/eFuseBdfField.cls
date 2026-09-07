VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eFuseBdfField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' From the table
Public msb As Long
Public LSB As Long
Private m_size As Long
Private m_hBytes As Long
Public Msb1st As Boolean

Public bankName As String
Public name As String
Public DefaulString As String
Public DefaultOrReal As DeaultReal
Public Default As String
Public BlowLocation As String
Public ReadOnly As Boolean

Public UsiLsb As Long
Public UsiMsb As Long
Public UsoLsb As Long
Public UsoMsb As Long
Public Llimit As Double
Public Hlimit As Double
Public BVLLimit As New SiteDouble
Public BVHLimit As New SiteDouble
Public xLlimit As String
Public xHlimit As String
Public LimitIsZero As Boolean
Public Resolution As Double
Public Algorithm As Algorithm
Public Description As String
Public IsEarlyFuse As Boolean
' EFuse Trim-Data
Private m_FieldDefault As New SiteVariant 'store the default values
Private m_TrimAteValue As New SiteVariant 'set it to be next trim dssc source
Private m_TrimmedData As New SiteVariant 'Identify as trim already, the set it to this variable
Private m_DsscValue As New SiteVariant 'This uses for storing the DSSC result in Hex format
Private m_DsscDecValue As New SiteDouble  'This uses for storing the DSSC result in Decimal format, only the field size <32bits which value could be refered
Private m_MeasureValue As New SiteDouble 'This uses for storing the IDS and VddBin values
Private m_FuseMeasureValue As New SiteDouble 'This uses for storing the IDS and VddBin values
Private m_TrimTestInstance As New SiteVariant
Private m_TrimAteDecValue As New SiteVariant ''202004xx for ap, add decimal data type for printing.

Private m_TrimFlagBySite As New SiteBoolean
Public Trimmed As Boolean
Public RealSetByDefault As Boolean
Public DefaultSetByUser As Boolean
Public m_PatTestPass As New SiteBoolean

Public crc_type As crctype  '20220602 Add for crc complement
Public SetWriteByBincut As Boolean

' Non-table data
Public DataRowNum As Long
Public myUtility As Utility

Private Sub Class_Initialize()
On Error GoTo errHandler
    Default = "0x0"
    TrimAteValue = "0x0"
    Msb1st = True
    UsiLsb = 0
    UsiMsb = 0
    UsoLsb = 0
    UsoMsb = 0
    Llimit = 0
    Hlimit = 0
    BVLLimit = 0
    BVHLimit = 0
    Resolution = 0
    FieldDefault = "0x0"
    DsscValue = "0x0"
    DsscDecValue = 0
    MeasureValue = 0
    Algorithm = alg_app
    ReadOnly = True
    DefaultOrReal = DeaultReal.dr_default
    m_TrimTestInstance = "-"
    m_TrimFlagBySite = False
    Trimmed = False
    IsEarlyFuse = False
    PatTestPass = True
    LimitIsZero = False
    RealSetByDefault = False
    DefaultSetByUser = False
    m_TrimAteDecValue = 0
    crc_type = crc_none   '20220602 Add for crc complement
    SetWriteByBincut = False
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "Class_Initialize")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Property Get TrimFlagBySite() As SiteVariant
    Set TrimFlagBySite = m_TrimFlagBySite
End Property

Public Property Get TrimAteDecValue() As SiteVariant
    Set TrimAteDecValue = m_TrimAteDecValue
End Property

Public Property Get size() As Long
    size = m_size
End Property

Public Property Let size(v As Long)
    m_size = v
    m_hBytes = (m_size - 1) \ 4 + 1
End Property

Public Property Get hBytes() As Long
    hBytes = m_hBytes
End Property

Public Property Get FieldDefault() As SiteVariant
    Set FieldDefault = m_FieldDefault
End Property

Public Property Get TrimAteValue() As SiteVariant
    Set TrimAteValue = m_TrimAteValue
End Property

Public Property Get TrimmedData() As SiteVariant
    Set TrimmedData = m_TrimmedData
End Property

Public Property Get DsscValue() As SiteVariant
    Set DsscValue = m_DsscValue
'    If Trimmed Then
'        m_TrimmedData = m_DsscValue 'Once set Dssc value, then update TrimmedData as well
'    End If
End Property

Public Property Get DsscDecValue() As SiteDouble
    Set DsscDecValue = m_DsscDecValue
End Property

Public Property Get MeasureValue() As SiteDouble
    Set MeasureValue = m_MeasureValue
End Property

Public Property Get FuseMeasureValue() As SiteDouble
    Set FuseMeasureValue = m_FuseMeasureValue
End Property

Public Property Get TrimTestInstance() As SiteVariant
    Set TrimTestInstance = m_TrimTestInstance
End Property

Public Property Get PatTestPass() As SiteVariant
    Set PatTestPass = m_PatTestPass
End Property

Public Sub InitVariables()
On Error GoTo errHandler
    m_TrimmedData = "0x0"
    m_TrimAteValue = "0x0"
    m_TrimAteDecValue = 0
    m_DsscValue = "0x0"
    m_DsscDecValue = 0
    m_PatTestPass = True
    m_MeasureValue = 0
    m_FuseMeasureValue = 0
    m_TrimTestInstance = "-"
    m_TrimFlagBySite = False
    Trimmed = False
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "InitVariables")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function IsOverflow(v As SiteVariant, Optional Force As Boolean = False, Optional bMsgBox As Boolean = True) As Boolean
On Error GoTo errHandler
Dim checkallsites As New SiteBoolean: checkallsites = False
Dim max1Loc As Long, binstr As String, site As Variant

    IsOverflow = True
        bMsgBox = False '220719 Prevent overflow message stop sorting
    If myUtility.IsHex(v) Then
         For Each site In theexec.sites.Selected
            binstr = myUtility.xHex2BinStr(v(site))
            max1Loc = InStr(1, binstr, "1")
            If max1Loc <> 0 Then
                max1Loc = Len(binstr) - InStr(1, binstr, "1") + 1
            End If
            If max1Loc > m_size Then
                If bMsgBox Then
                    myUtility.MessageBox "[Fatal Error] eFuse BDF Field is Overflow [" + name + "] @Site# " & site & " row " & CStr(DataRowNum) + " Maximum(Bits):" + CStr(m_size) + " Input(Bits):" + CStr(max1Loc)
                Else
                    Call Print_Error_Message(Error_Info, "eFuseBdfField", "IsOverflow", "eFuse BDF Field is Overflow [" + name + "] @Site# " & site & " row " & CStr(DataRowNum) + " Maximum(Bits):" + CStr(m_size) + " Input(Bits):" + CStr(max1Loc))
                End If
                checkallsites = True
                v(site) = 0
                gB_Fuse_Skip(site) = True
            End If
            IsOverflow = IsOverflow And checkallsites
        Next
    Else
        For Each site In theexec.sites.Selected
            If v(site) >= 2 ^ m_size Then
                If bMsgBox Then
                    myUtility.MessageBox "[Fatal Error] eFuse BDF Field is Overflow [" + name + "] @Site# " & site & " row " & CStr(DataRowNum) + " Maximum(Dec):" + CStr(2 ^ m_size) + " Input(Dec):" + CStr(v(site))
                Else
                    Call Print_Error_Message(Error_Info, "eFuseBdfField", "IsOverflow", "eFuse BDF Field is Overflow [" + name + "] @Site# " & site & " row " & CStr(DataRowNum) + " Maximum(Dec):" + CStr(2 ^ m_size) + " Input(Dec):" + CStr(v(site)))
                End If
                checkallsites = True
                v(site) = 0
                gB_Fuse_Skip(site) = True
            End If
            IsOverflow = IsOverflow And checkallsites
        Next
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "IsOverflow")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function xIsOverflow(v As String, Optional Force As Boolean = False, Optional bMsgBox As Boolean = True) As Boolean
On Error GoTo errHandler
Dim max1Loc As Long, binstr As String
Dim site As Variant 'Carter, 20240304''No use in the future??

    xIsOverflow = False
    If myUtility.xIsHex(v) Then
        If Not Force Then xIsOverflow = False: Exit Function ' skip it to save test time
        binstr = myUtility.xHex2BinStr(v)
        max1Loc = Len(binstr) - InStr(1, binstr, "1") + 1
        If max1Loc > m_size And max1Loc < Len(binstr) Then
            If bMsgBox Then
                myUtility.MessageBox "[Fatal Error] eFuse BDF Field is Overflow [" + name + "] @Site# " & site & " row " & DataRowNum
            Else
                Call Print_Error_Message(Error_Info, "eFuseBdfField", "xIsOverflow", "eFuse BDF Field is Overflow [" + name + "] @Site# " & site & " row " & DataRowNum)
            End If
            v = "0x0"
            xIsOverflow = True
        End If
    Else
        If CLng(v) >= 2 ^ m_size Then
             If bMsgBox Then
                 myUtility.MessageBox "[Fatal Error] eFuse BDF Field is Overflow [" + name + "] @Site# " & site & " row " & DataRowNum
             Else
                 Call Print_Error_Message(Error_Info, "eFuseBdfField", "xIsOverflow", "eFuse BDF Field is Overflow [" + name + "] @Site# " & site & " row " & DataRowNum)
             End If
             v = "0x0"
             xIsOverflow = True
        End If
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "xIsOverflow")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IdsValue_A() As Double
On Error GoTo errHandler

    IdsValue_A = (myUtility.Hex2Dbl(DsscValue) * Resolution) * 0.001  'to A

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "IdsValue_A")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function VddBinValue_V() As Double
On Error GoTo errHandler
Dim rt As Double

    rt = myUtility.Hex2Dbl(DsscValue)
    VddBinValue_V = IIf(rt <> 0, (myUtility.Hex2Dbl(DsscValue) * Resolution + BdfDataBase.BaseVoltage), 0)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "VddBinValue_V")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function BaseValue_V() As Double
On Error GoTo errHandler
Dim rt As Double

    rt = myUtility.Hex2Dbl(DsscValue)
    BaseValue_V = IIf(rt <> 0, ((myUtility.Hex2Dbl(DsscValue) + 1) * Resolution) * 0.001, 0) 'to V

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "BaseValue_V")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IdsValue_Hex(value As Double) As String
On Error GoTo errHandler
    IdsValue_Hex = myUtility.Dec2HexStr((value * 1000 / Resolution), size / 4 + 1)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "IdsValue_Hex")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function VddBinValue_Hex(value As Double) As String
On Error GoTo errHandler

    VddBinValue_Hex = IIf(value <> 0, myUtility.Dec2HexStr(((value - BdfDataBase.BaseVoltage) / Resolution), size / 4 + 1), 0)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "VddBinValue_Hex")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function BaseValue_Hex(value As Double) As String
On Error GoTo errHandler

    BaseValue_Hex = IIf(value <> 0, myUtility.Dec2HexStr(((value / Resolution) - 1), size / 4 + 1), 0) 'to V

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "BaseValue_Hex")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetFieldHLlimit(bankName As String)
On Error GoTo errHandler
Dim FusedTypeBank As eFuseBdfBank
Dim fusedField As eFuseBdfField
    
    Select Case UCase(bankName)
        Case "CMP":
            Set FusedTypeBank = BdfDataBase.Bank_Udr
        Case "CMPP", "CMP_P":
            Set FusedTypeBank = BdfDataBase.Bank_UdrP
        Case "CMPP0", "CMP_P0":
            Set FusedTypeBank = BdfDataBase.Bank_UdrP0
        Case "CMPP1", "CMP_P1":
            Set FusedTypeBank = BdfDataBase.Bank_UdrP1
        Case "CMPE", "CMP_E":
            Set FusedTypeBank = BdfDataBase.Bank_UdrE
        Case Else
    End Select
    Set fusedField = FusedTypeBank.Fields(name)
            ''set the field h/l limit of cmp bank  fromcorresponding bank
            ''   CMP - UDR
            '' CMPP - UDRP
            '' CMPE - UDRE
            
    ''20201217 modify bincut limit
    BVLLimit = fusedField.DsscDecValue
    BVLLimit = fusedField.DsscDecValue

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfField", "GetFieldHLlimit")
    If AbortTest Then Exit Function Else Resume Next
End Function
