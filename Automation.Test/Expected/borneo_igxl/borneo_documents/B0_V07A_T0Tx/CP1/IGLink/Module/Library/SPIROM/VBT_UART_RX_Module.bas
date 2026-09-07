Attribute VB_Name = "VBT_UART_RX_Module"
#Const isUFP = True
Option Explicit
Public Function UART_write_pmgr() As Long
On Error GoTo errHandler

    SendCmdOnly "pmgr bincut-check disable"

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_UART_RX_Module", "UART_write_pmgr")
    If AbortTest Then Exit Function Else Resume Next
End Function
'
'Public Function UART_write_run_sc_11() As Long
'
'    SendCmdOnly "sc run 11"
'
'End Function
