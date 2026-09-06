Attribute VB_Name = "VBT_UART_TX_Module"
Option Explicit



Public Function UART_read_response() As Long
On Error GoTo errHandler

    Dim i As Long
'    With thehdw.Protocol.ports("UART_PA").NWire.Frames("UART_Snd")
'        .Fields("Data_in").Value = 13   '10
'        .Execute
'    End With
    
    For i = 0 To 200
        With TheHdw.Protocol.ports("UART_TX").NWire.Frames("UART_Rcv")
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With
    Next i
  
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_UART_TX_Module", "UART_read_response")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UART_read_response_extended() As Long
On Error GoTo errHandler

    Dim i As Long
'    With thehdw.Protocol.ports("UART_PA").NWire.Frames("UART_Snd")
'        .Fields("Data_in").Value = 13   '10
'        .Execute
'    End With
    
    For i = 0 To 20000   '1500
        With TheHdw.Protocol.ports("UART_TX").NWire.Frames("UART_Rcv")
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With
    Next i
  
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_UART_TX_Module", "UART_read_response_extended")
    If AbortTest Then Exit Function Else Resume Next
End Function
'

Public Function UART_boot() As Long
On Error GoTo errHandler

    Dim i As Long
    For i = 0 To 5000 'TTR-750
        With TheHdw.Protocol.ports("UART_TX").NWire.Frames("UART_Rcv")
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With
    Next i
  
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_UART_TX_Module", "UART_boot")
    If AbortTest Then Exit Function Else Resume Next
End Function
