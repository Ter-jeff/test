Attribute VB_Name = "GPIBWrappers_NI"
'
' (c) Teradyne, Inc, 1997, 1998
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license agreement from Teradyne, Inc.
'
' Revision History:
' ----------------
'
' 10 Sep 2007 Jeffrey Tan: Initial release for ESI 21XX Laser Trimmer
'
'
Const NO_SECONDARY_ADDR = 0    ' Secondary address of device
Const EOTMODE = NLend          ' LF is sent with the EOI line asserted after the last byte of buffer
Const EOSMODE = XEOS Or REOS   ' XEOS = Send EOI with EOS byte; REOS = Terminate read on EOS
Public dev As Integer
Const OK = 1
Const Error = 2


Public Const SRQ_TJCAL_PASS = &HE3      ' TJ Cal Pass
Public Const SRQ_TJCAL_FAIL = &HE5      ' TJ Cal fail
Public Const SRQ_TCR_PASS = &HE4        ' TCR Pass
Public Const SRQ_TCR_FAIL = &HE9        ' TCR fail
Public Const SRQ_HDR_PASS = &HE4        ' HDR Pass
Public Const SRQ_HDR_FAIL = &HE9        ' HDR fail
Public Const SRQ_HFLOW_PASS = &H59      ' HFlow Pass
Public Const SRQ_TCS_PASS = &H59        ' TCS Pass

Public ATC_HDPlus_Mode As Integer
Public Const STANDARD_MODE = 6
Public Const HDPLUS_MODE = 1
Public Const ATCS_HTF_MODE = 2
'Public Const ATCS_DTF_MODE = 3
'Public Const ATCS_HTFPF_MODE = 4
'Public Const ATCS_DTFPF_MODE = 5
Public Const ATCS_HTF_PF_MODE = 3                       ' HTF with PF
Public Const ATCS_HTF_DieOffset_GPIB_MODE = 4           ' HTF with die offset by GPIB
Public Const ATCS_HTF_DieOffset_Temperature_MODE = 5    ' HTF with die offset by die temperature

Public GPIBCtrl As New Cls_HandleTemperature



Private Function AddIbcnt() As String
    AddIbcnt = Chr$(13) + Chr$(10) + "ibcnt = 0x" + Hex$(ibcnt)
End Function

Private Function AddIberr() As String
    If (ibsta And EERR) Then
        If (iberr = EDVR) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = EDVR <DOS Error>"
        If (iberr = ECIC) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = ECIC <Not CIC>"
        If (iberr = ENOL) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = ENOL <No Listener>"
        If (iberr = EADR) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = EADR <Address Error>"
        If (iberr = EARG) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = EARG <Invalid argument>"
        If (iberr = ESAC) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = ESAC <Not Sys Ctrlr>"
        If (iberr = EABO) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = EABO <Op. aborted>"
        If (iberr = ENEB) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = ENEB <No GPIB board>"
        If (iberr = EOIP) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = EOIP <Async I/O in prg>"
        If (iberr = ECAP) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = ECAP <No capability>"
        If (iberr = EFSO) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = EFSO <File sys. error>"
        If (iberr = EBUS) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = EBUS <Command error>"
        If (iberr = ESTB) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = ESTB <Status byte lost>"
        If (iberr = ESRQ) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = ESRQ <SRQ stuck high>"
        If (iberr = ETAB) Then AddIberr = Chr$(13) + Chr$(10) + "iberr = ETAB <Table overflow>"
    Else
        AddIberr = Chr$(13) + Chr$(10) + "iberr = " + str$(iberr)
    End If
End Function

Private Function AddIbsta() As String
    
    sta$ = Chr$(13) + Chr$(10) + "ibsta = &H" + Hex$(ibsta) + " <"
    If (ibsta And EERR) Then sta$ = sta$ + " ERR"
    If (ibsta And TIMO) Then sta$ = sta$ + " TIMO"
    If (ibsta And EEND) Then sta$ = sta$ + " END"
    If (ibsta And SRQI) Then sta$ = sta$ + " SRQI"
    If (ibsta And RQS) Then sta$ = sta$ + " RQS"
    If (ibsta And CMPL) Then sta$ = sta$ + " CMPL"
    If (ibsta And LOK) Then sta$ = sta$ + " LOK"
    If (ibsta And RREM) Then sta$ = sta$ + " REM"
    If (ibsta And CIC) Then sta$ = sta$ + " CIC"
    If (ibsta And AATN) Then sta$ = sta$ + " ATN"
    If (ibsta And TACS) Then sta$ = sta$ + " TACS"
    If (ibsta And LACS) Then sta$ = sta$ + " LACS"
    If (ibsta And DTAS) Then sta$ = sta$ + " DTAS"
    If (ibsta And DCAS) Then sta$ = sta$ + " DCAS"
    sta$ = sta$ + ">"
    AddIbsta = sta$
End Function

Private Sub GpibErr(msg$)
    msg$ = msg$ + AddIbsta() + AddIberr() + AddIbcnt() + Chr$(13) + Chr$(13) + "I'm quitting!"
    MsgBox msg$, vbOKOnly + vbExclamation, "Error"
    
    '  Take the device offline.
    
    ilonl dev, 0
    
    'End
End Sub

Public Function ConnectDevice_NI(ByVal boardNum As Integer, ByVal addr As Long, _
                                 ByVal TimeoutCode As Integer) As Boolean

On Error GoTo ConnectDeviceErr
    
    ConnectDevice_NI = True
    
'    Dev = ildev(boardNum, addr, NO_SECONDARY_ADDR, TimeoutCode, EOTMODE, EOSMODE)
   dev = ildev(boardNum, addr, NO_SECONDARY_ADDR, TimeoutCode, EOTMODE, 1034)   ' 3082
    If (ibsta And EERR) Then
        GpibErr ("Error opening device.")
        ConnectDevice_NI = False
    End If
    
'  Clear the internal or device functions of the device.  If the error bit
'  EERR is set in ibsta, call GpibErr with an error message.
    
    ilclr dev
    If (ibsta And EERR) Then
        GpibErr ("Error clearing device.")
        ConnectDevice_NI = False
    End If

    Exit Function

ConnectDeviceErr:

    'Add your own error processing here.
    MsgBox "Error in ConnectDevice_NI()"
    On Error GoTo 0

End Function


'This function frees device configuration for a the connected device
'
Public Function DisconnectDevice_NI() As Boolean

On Error GoTo DisconnectDeviceErr

    DisconnectDevice_NI = True

    ' Take the device offline.

    ilonl dev, 0
    If (ibsta And EERR) Then
        GpibErr ("Error putting device offline.")
        DisconnectDevice_NI = False
    End If

    Exit Function

DisconnectDeviceErr:

    'Add your own error processing here.
    MsgBox "Error in DisconnectDevice()"
    On Error GoTo 0

End Function

'This function sends a string (command) to the connected GPIB device.
'
Public Function GPIBSendCommand_NI(ByVal Cmd As String) As Long

On Error GoTo GPIBSendCommandErr
    
    GPIBSendCommand_NI = OK
    Dim terminator As String
    DoEvents
    'Sleep 2000
    
    terminator = Chr$(val("&H 0D")) + Chr$(val("&H 0A"))    ' CR+LF
    Cmd = Cmd + terminator
    ilwrt dev, Cmd, Len(Cmd)
    If (ibsta And EERR) Then
        GpibErr ("Error writing to device.")
        GPIBSendCommand_NI = Error
    End If
 
    Exit Function

GPIBSendCommandErr:

    ' Add your own error processing here.
    MsgBox "Error in GPIBSendCommand_NI() sending '" & Cmd & "'"
    On Error GoTo 0

End Function

'This function reads data from the GPIB device.
'
Public Function GPIBReadData_NI(ByRef ReplyBuf As String) As Long

On Error GoTo GPIBReadDataErr

    GPIBReadData_NI = OK
    
    ReplyBuf = Space$(200)
    ilrd dev, ReplyBuf, Len(ReplyBuf)
    If (ibsta And EERR) Then
        GpibErr ("Error reading from device.")
        GPIBReadData_NI = Error
    End If

    Exit Function

GPIBReadDataErr:

    ' Implement your error processing here.
    MsgBox "Error in GPIBReadData_NI()"
    On Error GoTo 0

End Function

'This function writes a command (that expects to return data) and then reads the reply.
'
Public Function GPIBSendQuery_NI(ByVal Cmd As String, ByRef ReplyBuf As String) As Long
    Dim Status As Long
    
    Status = Error
    
    If OK = GPIBSendCommand_NI(Cmd) Then Status = GPIBReadData_NI(ReplyBuf)
    GPIBSendQuery_NI = Status
    
End Function

'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
Public Function GPIB_RegKeyRead(i_RegKey As String) As String
    Dim myWS As Object
    
    On Error Resume Next
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'read key from registry
    GPIB_RegKeyRead = myWS.RegRead(i_RegKey)
    
End Function

'returns True if the registry key i_RegKey was found
'and False if not
Public Function RegKeyExists(i_RegKey As String) As Boolean
    Dim myWS As Object
    
    On Error GoTo ErrorHandler
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'try to read the registry key
    myWS.RegRead i_RegKey
    'key was found
    RegKeyExists = True
    Exit Function
    
ErrorHandler:
    'key was not found
    RegKeyExists = False
    
End Function



