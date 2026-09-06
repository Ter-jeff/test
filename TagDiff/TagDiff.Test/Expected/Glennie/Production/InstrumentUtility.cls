VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstrumentUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private DCVSpin() As String
Private dcviPin() As String
Private ioPin() As String
Private DCVSPins As String
Private DCVIPins As String
Private iopins As String
Private dcvsPinCnt As Long
Private dcviPinCnt As Long
Private ioPinCnt As Long
Private Const moduleNmae As String = "InstrumentUtility"
Private functionName As String
Sub Initialize(pins As String)
    SplitInstrument pins
End Sub
''' get property
Public Property Get GetDCVSPin() As String()
    GetDCVSPin = DCVSpin
End Property
Public Property Get GetDCVSPins() As String
    GetDCVSPins = DCVSPins
End Property
Public Property Get GetDCVIPin() As String()
    GetDCVIPin = dcviPin
End Property
Public Property Get GetDCVIPins() As String
    GetDCVIPins = DCVIPins
End Property
Public Property Get GetDCVSPinCount() As Long
    GetDCVSPinCount = dcvsPinCnt
End Property
Public Property Get GetDCVIPinCount() As Long
    GetDCVIPinCount = dcviPinCnt
End Property
Public Property Get GetIOPin() As String()
    GetIOPin = ioPin
End Property
Public Property Get GetIOPins() As String
    GetIOPins = iopins
End Property
Public Property Get GetIOPinCount() As Long
    GetIOPinCount = ioPinCnt
End Property

' [20230407][T-All] Add to classify pin to DCVS and DCVI
' [20231228][T-All][Tank] Add classify pin to Digital
' [20240223][T-All][Clyde] Add DCVS, DCVI and Digital pin count
Private Sub SplitInstrument(pins As String)
On Error GoTo errHandler
    Dim tempPinAry() As String
    Dim TempPinCnt As Long
    Dim sPinName As String
    Dim s_Msg As String
    Dim s_TempPin As String
    Dim i As Integer
    
    functionName = "SplitInstrument"
    dcvsPinCnt = 0
    dcviPinCnt = 0
    ioPinCnt = 0
    TheExec.DataManager.DecomposePinList pins, tempPinAry, TempPinCnt
    ReDim DCVSpin(TempPinCnt - 1) As String
    ReDim dcviPin(TempPinCnt - 1) As String
    ReDim ioPin(TempPinCnt - 1) As String
    
    For i = 0 To TempPinCnt - 1
        sPinName = LCase(tempPinAry(i))
        If gl_GetInstrumentType_Dic.Exists(sPinName) Then
            If LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvs*" Then
                DCVSPins = CombineStringList(DCVSPins, sPinName)
                DCVSpin(dcvsPinCnt) = tempPinAry(i)
                dcvsPinCnt = dcvsPinCnt + 1
            ElseIf LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvi*" Then
                DCVIPins = CombineStringList(DCVIPins, sPinName)
                dcviPin(dcviPinCnt) = tempPinAry(i)
                dcviPinCnt = dcviPinCnt + 1
            ElseIf LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*i/o*" Then
                iopins = CombineStringList(iopins, sPinName)
                ioPin(ioPinCnt) = tempPinAry(i)
                ioPinCnt = ioPinCnt + 1
            Else
                s_TempPin = CombineStringList(s_TempPin, sPinName)
            End If
        Else
            s_TempPin = CombineStringList(s_TempPin, sPinName)
        End If
    Next i

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, moudleName, functionName)
    If AbortTest Then Exit Sub Else Resume Next
End Sub


'[20230311][T-All][Clyde] Apply DCVI Power, Modularize code
Private Function ApplyDCVIPower(Optional ForceValue As Double = 0, Optional forceType As forceType = fVoltage, _
    Optional gateStatus As gateStatus, Optional connectStatus As ConnectedStatus)
On Error GoTo errHandler
    
    functionName = "ApplyDCVIPower"
    If forceType <> fCurrent Then
        With TheHdw.DCVI.pins(DCVIPins)
            If connectStatus = Connect Then .Connect tlDCVIConnectDefault
            If forceType = fVoltage Then .Voltage.value = ForceValue
            If gateStatus = gOn Then .Gate = True
            If gateStatus = gOff Then .Gate = False
            If connectStatus = Disconnect Then .Disconnect tlDCVIConnectDefault
        End With
    Else
        Call Print_Error_Message(Error_Info, moduleName, functionName, "This function not support force current")
    End If
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20230311][T-All][Clyde] Apply DCVS Power, Modularize code
Private Function ApplyDCVSPower(Optional ForceValue As Double = 0, Optional forceType As forceType = fVoltage, _
    Optional gateStatus As gateStatus, Optional connectStatus As ConnectedStatus, Optional applyOutput As DCVSoutput = dKeep, _
    Optional switchOutput As Boolean = False)
On Error GoTo errHandler

    functionName = "ApplyDCVSPower"
    If forceType <> fCurrent Then
        With TheHdw.DCVS.pins(DCVSPins)
            If connectStatus = Connect Then .Connect
            If forceType = fVoltage Then
                If applyOutput = dKeep Then
                    .Voltage.value = ForceValue
                ElseIf applyOutput = dVmain Then
                    .Voltage.Main.value = ForceValue
                    If switchOutput Then .Voltage.Output = tlDCVSVoltageMain
                Else
                    .Voltage.Alt.value = ForceValue
                    If switchOutput Then .Voltage.Output = tlDCVSVoltageAlt
                End If
            End If
            If gateStatus = gOn Then .Gate = True
            If gateStatus = gOff Then .Gate = False
            If connectStatus = Disconnect Then .Disconnect
        End With
    Else
        Call Print_Error_Message(Error_Info, moduleName, functionName, "This function not support force current")
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20230311][T-All][Clyde] Apply DCVS and DCVI Power, Modularize code
Public Function ApplyPower(Optional ForceValue As Double = 0, Optional forceType As forceType = fVoltage, _
    Optional gateStatus As gateStatus, Optional connectStatus As ConnectedStatus, Optional ByVal applyOutput As DCVSoutput = dKeep, _
    Optional switchOutput As Boolean = False)
On Error GoTo errHandler

    functionName = "ApplyPower"
    
    If DCVSPins <> "" Then
        ApplyDCVSPower ForceValue, forceType, gateStatus, connectStatus, applyOutput, switchOutput
    End If
    
    If DCVIPins <> "" Then
        ApplyDCVIPower ForceValue, forceType, gateStatus, connectStatus
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, moduleName, functionName)
    If AbortTest Then Exit Function Else Resume Next
End Function

