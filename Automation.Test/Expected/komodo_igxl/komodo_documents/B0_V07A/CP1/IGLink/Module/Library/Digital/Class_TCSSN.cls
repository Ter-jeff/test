VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_TCSSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TotalCycle As Long
Public TotalCoreCnt As Long
Public PinData As String
Public PinCnt As Long
Public dic_index2core As New Dictionary
Public dic_core2Index As New Dictionary
Private m_Packet() As Variant
Public PacketLength As Long
Public RepeatLength As Long
Private m_PacketMap() As Variant
Private m_RepeatMap() As Variant

Private Sub Class_Initialize()
    TotalCycle = 0
    TotalCoreCnt = 0
    PinData = vbNullString
    PinCnt = 0
    dic_index2core.RemoveAll
    dic_core2Index.RemoveAll
    ReDim m_Packet(0) As Variant
    ReDim m_PacketMap(0) As Variant
    ReDim m_RepeatMap(0, 0) As Variant
    PacketLength = 0
    RepeatLength = 0
End Sub

Public Property Let Packet(i As Long, value As Variant)
m_Packet(i) = value
End Property

Public Property Get Packet(i As Long) As Variant
Packet = m_Packet(i)
End Property

Public Property Let PacketMap(i As Long, value As Variant)
m_PacketMap(i) = value
End Property

Public Property Get PacketMap(i As Long) As Variant
PacketMap = m_PacketMap(i)
End Property

Public Property Let RepeatMap(i As Long, j As Long, value As Variant)
m_RepeatMap(i, j) = value
End Property

Public Property Get RepeatMap(i As Long, j As Long) As Variant
RepeatMap = m_RepeatMap(i, j)
End Property

Public Sub ResizePacketArr(Max_Index As Long)
    ReDim Preserve m_Packet(Max_Index) As Variant
End Sub

' Returns the Greatest Common Denominator (GCD) of the two numbers passed
' Ensure these are passed ByVal so they do not alter the original values of num1 and num2
Public Function GreatComDem(ByVal num1 As Long, ByVal num2 As Long) As Long
    Dim Remainder As Long
    While num2 <> 0:
        ' num1 and num2 are passed as Longs, so using VBA Mod should be fine
        Remainder = num1 Mod num2
        num1 = num2
        num2 = Remainder
    Wend
    GreatComDem = num1
End Function

' Generates a 1D array that serves as a map for what core each bit in a SSN packet corresponds to
Function GenPacketMap(BusWidth As Long)
    PacketLength = 0
    RepeatLength = 0
    
    ' Determine the number of bits in a packet by summing the given number of bits to each core in a packet
    Dim Inc As Long
    For Inc = LBound(m_Packet) To UBound(m_Packet)
        PacketLength = PacketLength + m_Packet(Inc)
    Next Inc
    
    ' Repeat Length is determined by the PacketLength divided by the Greates Common Denominator between the Bus Width and the Packet Length
    ' VBA does not provide a built-in GCD function, so a custom one is used
    RepeatLength = PacketLength / GreatComDem(BusWidth, PacketLength)
    ' Debug.Print "The pattern repeats every " & RepeatLength & " tester cycles"
    
    ReDim m_PacketMap(PacketLength - 1)
    
    Inc = 0
    Dim ArrayInc As Long
    ArrayInc = 0
    For Inc = LBound(m_Packet) To UBound(m_Packet)
        Dim Inc2 As Long
        For Inc2 = 1 To m_Packet(Inc)
            m_PacketMap(ArrayInc) = Inc
            ArrayInc = ArrayInc + 1
        Next Inc2
    Next Inc
    
    ' This block just prints the Packet Map to the datalog, should probably be removed
    ' TheExec.Datalog.WriteComment "Packet Map Pattern:"
    ' Dim Message As String
    ' For Inc = 0 To PacketLength - 1
    '     Message = Message & PacketMap(Inc) & " "
    ' Next Inc
    ' TheExec.Datalog.WriteComment Message
End Function

' Creates a map of the packet-shifting repeat pattern
' Function returns Variant so a 2D Array can be returned
Function GetBusRepeatMap(BusWidth As Long)
    Dim ArrayInc As Long
    Dim Inc As Long
    Dim Inc2 As Long
    Dim PacketLength As Long
    ' Get packet length by determining how long the packet map is
    PacketLength = UBound(m_PacketMap) + 1
    
    ' Size of the Repeat Map is determine by Bus Width and # of Cycles it takes for packet shift pattern to repeat
    ReDim m_RepeatMap(RepeatLength - 1, BusWidth - 1)
    
    ' ArrayInc tracks place in PacketMap
    ArrayInc = 0
    For Inc = 0 To RepeatLength - 1
        For Inc2 = 0 To BusWidth - 1
            m_RepeatMap(Inc, Inc2) = m_PacketMap(ArrayInc)
            ArrayInc = ArrayInc + 1
            If ArrayInc = PacketLength Then ArrayInc = 0
        Next Inc2
    Next Inc
    
    ' This block just prints the Repeat Map to the datalog, should probably be removed
    ' Dim Message As String
    ' TheExec.Datalog.WriteComment "Repeat Map Pattern:"
    ' For Inc = 0 To RepeatLength - 1
    '     Message = ""
    '     For Inc2 = 0 To BusWidth - 1
    '         Message = Message & RepeatMap(Inc, Inc2) & " "
    '     Next Inc2
    '     TheExec.Datalog.WriteComment Message
    ' Next Inc
End Function
