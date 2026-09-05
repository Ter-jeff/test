VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_HandleTemperature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ########################################################################################################
' =======================================================================================================
' To use Thermal Control with TEL prober
' =======================================================================================================
' 1. Import .bas files(NIGlobal.bas / Vbib_64.bas / GPIBWrappers_NI.bas and Cls_HandleTemperature.cls)
' 2. Make sure GPIBDeviceModel is "TELP8"
' 3. Add below 3 lines at OnProgramStarted()
'       If ATC_HDPlus_Mode <= 0 Then ATC_HDPlus_Mode = GPIBCtrl.Get_ATCS_HDPlus_Mode    ' To get temperature contorl mode type
'       If GPIBCtrl.TJCal_HDPlus = False Then Call MsgBox("TJ Caribration failed.", vbCritical)
'       If GPIBCtrl.TCR = False Then Call MsgBox("Thermal Control Run(TCR) failed.", vbCritical)
' 4. Add below 2 lines at OnProgramEnded()
'       Call GPIBCtrl.TCS                                                               ' To send STOP temperature contorl(TCS)
'       GPIBCtrl.GPIB_Disconnect                                                        ' To disconnect GPIB
' 5. Start contol the temperature by using one of below functions
'       Call GPIBCtrl.HDPlus_FlowControl(flow As String)
'       Call GPIBCtrl.HTF_ThermalControl_OffsetTemperature(TemperatureOffset As String)
'
'
' =======================================================================================================
' To use Thermal Control with Hontech handler
' =======================================================================================================
' 1. Import .bas files(NIGlobal.bas / Vbib_64.bas / GPIBWrappers_NI.bas / Cls_HandleTemperature.cls and VBT_FT_GPIB.bas)
' 2. Make sure GPIBDeviceModel is "SeikoEpsonNS5000"
' 3. Add below at OnProgramStarted()
'       Call GPIBCtrl.Set_Handler_VI(True)
' 4. Add below at OnProgramEnded()
'       Call GPIBCtrl.GPIB_Disconnect
'
'
' ########################################################################################################
Option Explicit

Const OK = 1
Const Error = 2

Private site As Variant
Private connectted_flag As Boolean
Private temperature_offset_list As String

Private Const GPIBDeviceModel = "TELP8"
'Private Const GPIBDeviceModel = "SeikoEpsonNS5000"



Public Function GPIB_Connect() As Boolean

    Dim RegPath As String
    Dim RegGPIBAdd As Long
    Dim result As Boolean
    
    ' To get driver GPIB address
    RegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Teradyne\IG-XL\Handler Drivers\" & GPIBDeviceModel & "\GPIBaddr"
    If is_64bitExcle() = True Then RegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Teradyne\IG-XL\Handler Drivers\" & GPIBDeviceModel & "\GPIBaddr"
    RegGPIBAdd = GPIB_RegKeyRead(RegPath)
    
    ' To connect GPIB
    result = ConnectDevice_NI(0, RegGPIBAdd, T100s)
    
    If result = True Then
        connectted_flag = True
    Else
        TheExec.Datalog.WriteComment "ERROR: GPIB_Connect Fail!!!"
        'TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
    End If
    
    GPIB_Connect = result

            
End Function


Public Function GPIB_Disconnect() As Boolean
    Dim result As Boolean

    ' To disconnect GPIB
    result = DisconnectDevice_NI()
    
    If result = True Then connectted_flag = False
    GPIB_Disconnect = result
End Function

Public Sub GPIB_ReadSRQ(ByVal ud As Integer, ByRef result As Integer)
    On Error GoTo GPIBReadSRQ_NIErr
    DoEvents
    
    result = 0
    
    Do While result = 0
        Call ibrsp(ud, result)
        Sleep 50
'        Call WaitSRQ(Dev, result)
    Exit Do
    Loop
    
    Exit Sub
    
GPIBReadSRQ_NIErr:
    Call MsgBox("Error in GPIBReadSRQ_NI()", vbCritical)
    On Error GoTo 0
    
End Sub

'##################################################################
' For Hontech handler to enable VI mode
'##################################################################
Public Function Set_Handler_VI(OnOff As Boolean)
    Dim Cmd As String:: Cmd = "POWERFOLLOWING"
    Dim parameter As String:: parameter = " _0"
    Dim reply As String
    Dim result As Boolean:: result = False
    
    If OnOff = True Then parameter = " _1"
    Cmd = Cmd & parameter
    
    ' When the GPIB is not conneted, then reconnect
    If connectted_flag = False Then Call GPIB_Connect
    
    ' Send command
    If GPIBSendQuery_NI(Cmd, reply) = Error Then
        Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
        
    Else
        If Trim(reply) <> "SETTINGOK" Then
            Call MsgBox("Handler replied setting fail(" & reply & ")", vbCritical)
        Else
            result = True
        End If
    End If
        
    Set_Handler_VI = result
End Function

'##################################################################
' For Hontech handler to change handler temperature offset value
'##################################################################
Public Sub Set_Handler_TemperatureOffset(offset As String, Optional site_num As Integer = "-1")
    Dim Cmd As String:: Cmd = "DEVICEOFFSET_"
    Dim reply As String
    Dim result As Boolean:: result = False
    Dim max_sites As Integer:: max_sites = TheExec.sites.Active.Count - 1
    Dim send_flag As Boolean:: send_flag = False
    
    ' When the GPIB is not conneted, then reconnect
    If connectted_flag = False Then Call GPIB_Connect
    
    ' To process parameter
    If site_num = -1 Then                                                   ' For sending whole offset list
        Cmd = Cmd & offset
        send_flag = True
    Else                                                                    ' For each site offset
        If site_num = 0 Then temperature_offset_list = ""
        
        temperature_offset_list = temperature_offset_list & offset & ","
        
        If max_sites = site_num Then
            Cmd = Cmd & temperature_offset_list
            send_flag = True
        End If
    End If
    
    If send_flag = True Then
        ' Send command
        If GPIBSendQuery_NI(Cmd, reply) = Error Then
            Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
            
        Else
            If Trim(reply) <> "SETTINGOK" Then
                Call MsgBox("Handler replied setting fail(" & reply & ")", vbCritical)
            Else
                result = True
            End If
        End If
    End If

End Sub

'##################################################################
' For Hontech handler to setup PID
'##################################################################
Public Function Set_PID(ByRef pid() As String) As Boolean
On Error GoTo errHandler
    Dim result As Boolean
    Dim i As Integer
    Dim pid_list As String
    Dim Cmd As String:: Cmd = "SETPID_"
    
    ' When the GPIB is not conneted, then reconnect
    If connectted_flag = False Then Call GPIB_Connect
    
    For i = 0 To UBound(pid)
        pid_list = pid_list & "," & pid(i)
    Next i
    Cmd = Cmd & right(pid_list, Len(pid_list) - 1) & ","

    If GPIBSendCommand_NI(Cmd) = Error Then
        Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
    Else
        result = True
    End If

    Set_PID = result
    Exit Function
        
errHandler:
    HandleExecIPError "Get_ATCS_HDPlus_Mode"
End Function



'##################################################################
' For TEL prober:
'##################################################################
Public Function TJCal_HDPlus() As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    Dim result As Boolean
    Dim received_SRQ As Integer
    
    result = False
    TheExec.Datalog.WriteComment "Current Test Instance : = " & TheExec.DataManager.instancename
    ' When the GPIB is not conneted, then reconnect
    If connectted_flag = False Then Call GPIB_Connect
    
    If ATC_HDPlus_Mode = HDPLUS_MODE Then                                           ' Mode type: HD+
        ' Send command
        Cmd = "HDR"
        If GPIBSendCommand_NI(Cmd) = Error Then
            Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
            
        Else
            Sleep (2000) '' Wait 2secs for readSRQ
             ' Read SRQ
            Call GPIB_ReadSRQ(dev, received_SRQ)
            If received_SRQ = SRQ_HDR_PASS Then
                result = True
            ElseIf received_SRQ = SRQ_HDR_FAIL Then
                Call MsgBox("TJCal_HDPlus(), Received SRQ: 0x" & CStr(Hex(received_SRQ)) & " HDR failed.", vbCritical)
            Else
                Call MsgBox("TJCal_HDPlus(), Received SRQ: 0x" & CStr(Hex(received_SRQ)), vbCritical)
            End If
        End If
    
   ElseIf ATC_HDPlus_Mode = ATCS_HTF_MODE Or _
        ATC_HDPlus_Mode = ATCS_HTF_PF_MODE Or _
        ATC_HDPlus_Mode = ATCS_HTF_DieOffset_GPIB_MODE Or _
        ATC_HDPlus_Mode = ATCS_HTF_DieOffset_Temperature_MODE Then                     ' Mode type: ATCS_HTF / HTF with die offset by GPIB
        ' Send command
        Cmd = "TJCAL"
        If GPIBSendCommand_NI(Cmd) = Error Then
            Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
            
        Else
            TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & " ; TCJCAL Command "
            Sleep (2000) '' Wait 2secs for readSRQ
             ' Read SRQ
            Call GPIB_ReadSRQ(dev, received_SRQ)
            TheExec.Datalog.WriteComment "TJCAL GPIB_ReadSRQ = " & received_SRQ
            

            If received_SRQ = SRQ_TJCAL_PASS Then
                result = True
            ElseIf received_SRQ = SRQ_HDR_FAIL Then
                Call MsgBox("TJCal_HDPlus(), Received SRQ: 0x" & CStr(Hex(received_SRQ)) & " TJCAL failed.", vbCritical)
            Else
                Call MsgBox("TJCal_HDPlus(), Received SRQ: 0x" & CStr(Hex(received_SRQ)), vbCritical)
            End If
            
        End If
        
    Else
        Call MsgBox("Unknow Mode: " & CStr(ATC_HDPlus_Mode), vbCritical)
        
    End If

    TJCal_HDPlus = result
    
    Exit Function
        
errHandler:
    HandleExecIPError "TJCal_HDPlus"
        
End Function

'##################################################################
' For TEL prober:
' Thermal Control Run(Start)
' When TCR pass, it return true
'##################################################################
Public Function TCR() As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    Cmd = "TCR"
    Dim result As Boolean
    Dim received_SRQ As Integer
    
    result = False
    TheExec.Datalog.WriteComment "Current Test Instance : = " & TheExec.DataManager.instancename
   If ATC_HDPlus_Mode = ATCS_HTF_MODE Or _
        ATC_HDPlus_Mode = ATCS_HTF_PF_MODE Or _
        ATC_HDPlus_Mode = ATCS_HTF_DieOffset_GPIB_MODE Or _
        ATC_HDPlus_Mode = ATCS_HTF_DieOffset_Temperature_MODE Then
        ' When the GPIB is not conneted, then reconnect
        If connectted_flag = False Then Call GPIB_Connect

        ' Send command
        If GPIBSendCommand_NI(Cmd) = Error Then
            Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
            
        Else
            TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & " ; TCR Command "
            Sleep (2000) '' Wait 2secs for readSRQ
             ' Read SRQ
            Call GPIB_ReadSRQ(dev, received_SRQ)
            TheExec.Datalog.WriteComment "TCR GPIB_ReadSRQ = " & received_SRQ
            
            If received_SRQ = SRQ_TCR_PASS Then
                result = True
            Else
                Call MsgBox("TCR(), Received SRQ: 0x" & CStr(Hex(received_SRQ)), vbCritical)
            End If
        End If
    Else
        Call MsgBox("Current is HD+ mode. Not able to use ATCS command(" & Cmd & ")", vbCritical)
    End If

    TCR = result
    
    Exit Function
        
errHandler:
    HandleExecIPError "TCR"
    
End Function

'##################################################################
' For TEL prober:
' Thermal Control Stop(Stop)
' When TCS pass, it return true
'##################################################################
Public Function TCS() As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    Cmd = "TCS"
    Dim result As Boolean
    Dim received_SRQ As Integer
    
    result = False
    TheExec.Datalog.WriteComment "Current is TCS"
    ' When the GPIB is not conneted, then reconnect
    If connectted_flag = False Then Call GPIB_Connect
    
    ' Send command
    If GPIBSendCommand_NI(Cmd) = Error Then
        Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
        
    Else
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & " ; TCS Command "
        Sleep (2000) '' Wait 2secs for readSRQ
        ' Read SRQ
        Call GPIB_ReadSRQ(dev, received_SRQ)
        TheExec.Datalog.WriteComment "TCS GPIB_ReadSRQ = " & received_SRQ
        
        If received_SRQ = SRQ_TCS_PASS Then
            result = True
        Else
            Call MsgBox("TCS(), Received SRQ: 0x" & CStr(Hex(received_SRQ)), vbCritical)
        End If
            
    End If

    TCS = result
    
    Exit Function
        
errHandler:
    HandleExecIPError "TCS"
End Function

'##################################################################
' For TEL prober:
' HD+ Flow Control
' When flow control pass, it return true
'##################################################################
Public Function HDPlus_FlowControl(flow As String) As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    Cmd = "HFLOW" & Format(flow, "000.0")
    Dim result As Boolean
    Dim received_SRQ As Integer
    
    result = False
    
    If ATC_HDPlus_Mode = HDPLUS_MODE Then
        ' When the GPIB is not conneted, then reconnect
        If connectted_flag = False Then Call GPIB_Connect
        
        ' Send command
        If GPIBSendCommand_NI(Cmd) = Error Then
            Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
            
        Else
            ' Read SRQ
            Call GPIB_ReadSRQ(dev, received_SRQ)
            If received_SRQ = SRQ_HFLOW_PASS Then
                result = True
            Else
                Call MsgBox("HDPlus_FlowControl(), Received SRQ: 0x" & CStr(Hex(received_SRQ)), vbCritical)
            End If
            
        End If

    Else
        Call MsgBox("Current is ATC mode. Not able to use HD+ command(" & Cmd & ")", vbCritical)
    End If
    
    HDPlus_FlowControl = result
    Exit Function
        
errHandler:
    HandleExecIPError "HDPlus_FlowControl"
End Function

'##################################################################
' For TEL prober:
' HD+ Flow Control
'##################################################################
Public Function HDPlus_RequestFlowValve(ByRef flowvalve As String) As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    Cmd = "?HFLOW"
    Dim result As Boolean
    Dim received_SRQ As Integer
    Dim ReplyBuf As String
    
    result = False
    
    If ATC_HDPlus_Mode = HDPLUS_MODE Then
        ' When the GPIB is not conneted, then reconnect
        If connectted_flag = False Then Call GPIB_Connect
        
        ' Send command
        If GPIBSendQuery_NI(Cmd, ReplyBuf) = Error Then
            Call MsgBox("Not able to send out request command(" & Cmd & ")", vbCritical)
            
        Else
            result = True
            flowvalve = Trim(Replace(ReplyBuf, vbCrLf, ""))

        End If

    Else
        Call MsgBox("Current is ATC mode. Not able to use HD+ command(" & Cmd & ")", vbCritical)
    End If
    
    HDPlus_RequestFlowValve = result
    Exit Function
        
errHandler:
    HandleExecIPError "HDPlus_RequestFlowValve"
End Function


'##################################################################
' For TEL prober:
' HTF Flow Control
' When temperature offset control pass, it return true
'##################################################################
Public Function HTF_ThermalControl_OffsetTemperature(TemperatureOffset As String) As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    ''20211129 Dim sign As String:: sign = "+"
    Dim result As Boolean
    Dim received_SRQ As Integer
   
    ''20211129 If TemperatureOffset < 0 Then sign = "-"
    ''20211129 cmd = "TCOFS" & sign & Format(TemperatureOffset, "00.0")
    Cmd = "TCOFS" & Format(TemperatureOffset, "00.0")
    result = False
    
    If ATC_HDPlus_Mode = ATCS_HTF_DieOffset_Temperature_MODE Or _
        ATC_HDPlus_Mode = ATCS_HTF_DieOffset_GPIB_MODE Then
        ' When the GPIB is not conneted, then reconnect
        If connectted_flag = False Then Call GPIB_Connect
        
        ' Send command
        If GPIBSendCommand_NI(Cmd) = Error Then
            Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
            
        Else
            TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & ";Temp offset " & Cmd
            Sleep (200) '' Wait 2secs for readSRQ
            ' Read SRQ
            Call GPIB_ReadSRQ(dev, received_SRQ)
            If received_SRQ = SRQ_HFLOW_PASS Then   ' Prober sends out 0x59 for command received pass and fail.
                result = True
            Else
                Call MsgBox("HTF_ThermalControl_OffsetTemperature(), Received SRQ: 0x" & CStr(Hex(received_SRQ)), vbCritical)
            End If
            
        End If

    Else
        Call MsgBox("Current prober is not " & Chr(34) & "HTF with die offset by die temp" & Chr(34) & _
                    " mode. So VBT is not able to send command" & vbCrLf & "(" & Cmd & ")", vbCritical)
    End If
    
    HTF_ThermalControl_OffsetTemperature = result
    Exit Function
        
errHandler:
    HandleExecIPError "HTF_ThermalControl_OffsetTemperature"
End Function



Private Function is_64bitExcle() As Boolean

    Dim RegPath As String
    Dim result As Boolean
    Dim temp As String
    
    RegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook\Bitness"
    temp = GPIB_RegKeyRead(RegPath)
    
    If temp = "x64" Then result = True
    
    is_64bitExcle = result
End Function
'##################################################################
' For TEL prober:
' TJVOL Flow Control Feedback Voltage
'##################################################################
Public Function TJVOL() As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    Dim result As Boolean
    Dim received As Double
    Dim reply As String
    Dim reply_temp As String
    Dim reply_Type As String
    Dim i As Integer
    
    
    TheExec.Datalog.WriteComment "Current Test Instance : = " & TheExec.DataManager.instancename
   
    Cmd = "TJVOL"
    If ATC_HDPlus_Mode = ATCS_HTF_DieOffset_Temperature_MODE Then
        ' When the GPIB is not conneted, then reconnect
        If connectted_flag = False Then Call GPIB_Connect
        
        ' Send command
        
        If GPIBSendQuery_NI(Cmd, reply) = Error Then
        
           Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
        
        Else
        
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & ";Temp offset " & Cmd

        reply_temp = reply
        For i = 0 To 7
            reply = mid(reply_temp, (i * 8) + 1, 8)
            received = Format(reply, "0.00000")
            TheExec.Datalog.WriteComment "TJVOL GPIB_ReadVoltage_" & i + 1 & " = " & received
        Next i
        
        End If
        

    Else
        Call MsgBox("Current prober is not HTF with die offset by temperature mode. So VBT is not able to send command" & vbCrLf & "(" & Cmd & ")", vbCritical)
    End If

    
    Exit Function
        
errHandler:
    HandleExecIPError "HTF_ThermalControl_OffsetTemperature"
End Function
'##################################################################
' For TEL prober:
' TJSLOPE Flow Control Feedback Temperature
'##################################################################
Public Function TJSLOPE(ChannelNum As String, TemperatureSlope As String, TemperatureOffset As String) As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    ''20211129 Dim sign As String:: sign = "+"
    Dim result As Boolean
    Dim received As Double
    Dim reply As String
    
    TheExec.Datalog.WriteComment "Current Test Instance : = " & TheExec.DataManager.instancename
    
    TemperatureSlope = Format(TemperatureSlope, "0000.00")
    TemperatureOffset = Format(TemperatureOffset, "0000.00")
    
    If CDbl(TemperatureSlope) >= 0 Then
        TemperatureSlope = "+" & TemperatureSlope
    End If
    
    If CDbl(TemperatureOffset) >= 0 Then
        TemperatureOffset = "+" & TemperatureOffset
    End If
    
    Cmd = "TJSLOPE" & ChannelNum & TemperatureSlope & TemperatureOffset
    result = False
    
    If ATC_HDPlus_Mode = ATCS_HTF_DieOffset_Temperature_MODE Then
        ' When the GPIB is not conneted, then reconnect
        If connectted_flag = False Then Call GPIB_Connect
        
        ' Send command
        
        If GPIBSendQuery_NI(Cmd, reply) = Error Then
        
        Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
        
        Else
        
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & ";Temp offset " & Cmd

       ' TheExec.Datalog.WriteComment "TJSLOPE GPIB_ReadTemperature = " & reply
        received = Format(reply, "000.0")
        ATC_TJ_measurement = received '''''''''''''HS 220511
        TheExec.Datalog.WriteComment "TJSLOPE GPIB_ReadTemperature = " & received
        
        End If

      

    Else
        Call MsgBox("Current prober is not HTF with die offset by temperature mode. So VBT is not able to send command" & vbCrLf & "(" & Cmd & ")", vbCritical)
    End If

    
    Exit Function
        
errHandler:
    HandleExecIPError "HTF_ThermalControl_OffsetTemperature"
End Function
'##################################################################
' For TEL prober:
' TJCH Flow Control Temperature For Channel Used
'##################################################################
Public Function TJCH(ChannelCnt As String, ChannelList As String) As Boolean
On Error GoTo errHandler

    Dim Cmd As String
    Dim result As Boolean
    Dim received_SRQ As Integer
    Dim channelStr As String
    
    TheExec.Datalog.WriteComment "Current Test Instance : = " & TheExec.DataManager.instancename

    channelStr = Replace(ChannelList, ",", "")
    

    Cmd = "TJCH" & ChannelCnt & channelStr
    result = False
    
    If ATC_HDPlus_Mode = ATCS_HTF_DieOffset_Temperature_MODE Then
        ' When the GPIB is not conneted, then reconnect
        If connectted_flag = False Then Call GPIB_Connect
        
        ' Send command
        
         If GPIBSendCommand_NI(Cmd) = Error Then
            Call MsgBox("Not able to send out command(" & Cmd & ")", vbCritical)
            
         Else
            TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & ";Temp offset " & Cmd
            Sleep (2000) '' Wait 2secs for readSRQ
            ' Read SRQ
            Call GPIB_ReadSRQ(dev, received_SRQ)
            TheExec.Datalog.WriteComment "TJCH GPIB_ReadSRQ = " & received_SRQ
            
            If received_SRQ = SRQ_HFLOW_PASS Then
                result = True
            
            Else
                Call MsgBox("TJCH(), Received SRQ: 0x" & CStr(Hex(received_SRQ)), vbCritical)
            End If
                 
                                  
         End If
      TJCH = result

    Else
        Call MsgBox("Current prober is not HTF with die offset by temperature mode. So VBT is not able to send command" & vbCrLf & "(" & Cmd & ")", vbCritical)
    End If

    
    TJCH = result
    Exit Function
        
errHandler:
    HandleExecIPError "HTF_ThermalControl_OffsetTemperature"
End Function
