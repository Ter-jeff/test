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
    
    For i = 0 To 3000   '1500
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
    Dim reftime As Double
    Dim Elapsetime As Double
    Dim Char_Length As Integer
    
    'Flag is defined in OnProgramValidated to separater the CORR and PROD
    If RTOS_Is_PROD = True Then
        Char_Length = 20000
    Else
        Char_Length = 20000
    End If
    
        
    'reftime = Theexec.Timer
    For i = 0 To Char_Length 'TTR-1200 for CORR, TTR-1050 for PROD for B1 binary
        With TheHdw.Protocol.ports("UART_TX").NWire.Frames("UART_Rcv")
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With
    Next i

    'Elapsetime = Theexec.Timer(reftime)
  
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_UART_TX_Module", "UART_boot")
    If AbortTest Then Exit Function Else Resume Next
End Function
