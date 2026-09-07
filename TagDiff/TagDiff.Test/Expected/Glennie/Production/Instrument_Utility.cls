VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Instrument_Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const COREPOWER = "All_Power"
Private DCVS_Pins_() As String
Private DCVI_Pins_() As String
Private DCVS_PinCnt_ As Long
Private DCVI_PinCnt_ As Long

Sub Class_Initialize()
    SplitInstrumentOfPowerPins COREPOWER
End Sub
''' get property
Public Property Get GetDCVSPinsFromCorePower() As String()
    GetDCVSPinsFromCorePower = DCVS_Pins_
End Property
Public Property Get GetDCVIPinsFromCorePower() As String()
    GetDCVIPinsFromCorePower = DCVI_Pins_
End Property
Public Property Get GetDCVSPinCountFromCorePower() As Long
    GetDCVSPinCountFromCorePower = DCVS_PinCnt_
End Property
Public Property Get GetDCVIPinCountFromCorePower() As Long
    GetDCVIPinCountFromCorePower = DCVI_PinCnt_
End Property

Private Sub SplitInstrumentOfPowerPins(pins As String)
On Error GoTo errHandler
    Dim pin_grp() As String
    Dim Pin_Cnt As Long
    Dim pin_type As String
    Dim pin As Variant
    
    TheExec.DataManager.DecomposePinList pins, pin_grp, Pin_Cnt
    If Pin_Cnt = 1 Then
        If UCase(TheExec.DataManager.ChannelType(pins)) = "N/C" Then Exit Sub
        pin_type = UCase(GetInstrument(pins, 0))
        Select Case pin_type
            Case "DC-07"
                appendToArray DCVI_Pins_, pins
                DCVI_PinCnt_ = DCVI_PinCnt_ + 1
            Case "VHDVS", "HEXVS", "VSM", "VS-5A", "VS-800MA"
                appendToArray DCVS_Pins_, pins
                DCVS_PinCnt_ = DCVS_PinCnt_ + 1
        End Select
    Else
        For Each pin In pin_grp
            SplitInstrumentOfPowerPins CStr(pin)
        Next pin
    End If

    Exit Sub
errHandler:
    Stop
    Resume Next
End Sub

