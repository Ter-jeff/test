Attribute VB_Name = "VBT_GPIB"
Option Explicit

Public START_DATE_TIME As String

Public qqq As Integer
Public Tj_Tester  As Double
Public Tj_Tester_arr(4) As Double
Public TJ_Offset_Cal As Double
Public TJ_Slope_Cal As Double

Public Tj_measurement_pin As String
Public Temperature_junction As Double
Public TJ_target As Double
Public ATC_TJ_measurement As Double
Public TJVOL_measurement As Double

Public Function ATC_ReadTemp() As Long

    Dim CommandBuf As String
    Dim ReplyBuf As String
    Dim StartTime As Double
    Dim endTime As Double
    
    CommandBuf = "TEMPARM?"

    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Debug.Print Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    Debug.Print Trim(ReplyBuf)

End Function

Public Function ATC_ControlTemp(temp As String) As Long

    Dim CommandBuf As String
    Dim ReplyBuf As String
    Dim StartTime As Double
    Dim endTime As Double

    CommandBuf = "DEVICETEMP_" & temp
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Debug.Print Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    Debug.Print Trim(ReplyBuf)
    
End Function

Public Function ATC_GetPowerTableParameter(temp As String) As Long

    Dim CommandBuf As String
    Dim ReplyBuf As String
    
    If temp > 19 And temp < 151 Then
        CommandBuf = "GETPFCPARAMETER_" & temp
        TheExec.Datalog.WriteComment Trim(CommandBuf)
        Debug.Print Trim(CommandBuf)
        Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
        TheExec.Datalog.WriteComment Trim(ReplyBuf)
        Debug.Print Trim(ReplyBuf)
    Else
        TheExec.Datalog.WriteComment "Your input temp is out of spec!"
    End If

End Function

Public Function Get_Voltage() As Long

    '''''''''Use to get Voltage
    
    Dim CommandBuf As String
    Dim ReplyBuf As String
    
    CommandBuf = "GETVOLTAGE"
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Debug.Print Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    Debug.Print Trim(ReplyBuf)
    
End Function

Public Function Read_Tj() As Long

    '''''''''Use to get Tj temperature
    
    Dim CommandBuf As String
    Dim ReplyBuf As String
    
    CommandBuf = "READTJ"
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Debug.Print Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    Debug.Print Trim(ReplyBuf)
    
End Function

Public Function Get_SlopeOffset() As Long

    '''''''''Use to get slope and offset
    
    Dim CommandBuf As String
    Dim ReplyBuf As String
    
    CommandBuf = "GETSLOPEOFFSET"
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Debug.Print Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    Debug.Print Trim(ReplyBuf)
    
End Function

Public Function ATC_SetControlMode(mode As String) As Long

    ''''''''''0 : Tc Mode
    ''''''''''1 : Tj Mode
    
    Dim CommandBuf As String
    Dim ReplyBuf As String
    
    If Not right(mode, 1) = ";" Then
        mode = mode & ";"
    End If
    
    'If left(Mode, 1) = 0 Or left(Mode, 1) = 1 Or left(Mode, 1) = 2 Then
    If left(mode, 1) = 0 Or left(mode, 1) = 1 Then
        CommandBuf = "ATCCONTROLMODE_" & mode
        TheExec.Datalog.WriteComment Trim(CommandBuf)
        Debug.Print Trim(CommandBuf)
        Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
        TheExec.Datalog.WriteComment Trim(ReplyBuf)
        Debug.Print Trim(ReplyBuf)
    Else
        TheExec.Datalog.WriteComment "Your input mode doesn't exist!"
    End If
    
End Function

Public Function Set_SlopeOffset(SlopeOffset As String) As Long

    Dim CommandBuf As String
    Dim ReplyBuf As String
    Dim TempSplit() As String
    
    TempSplit = Split(SlopeOffset, ",")
    
    If Not UBound(TempSplit) Mod 2 = 1 Then
        TheExec.Datalog.WriteComment "Please check your slope and offset string!"
        Exit Function
    End If
    
    If Not right(SlopeOffset, 1) = ";" Then
        SlopeOffset = SlopeOffset & ";"
    End If
    
    CommandBuf = "SETSLOPEOFFSET_" & SlopeOffset
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Debug.Print Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    Debug.Print Trim(ReplyBuf)
    
End Function

Public Function ATC_SetPowerTableParameter(temp As String, Watt As String) As Long

    Dim CommandBuf As String
    Dim ReplyBuf As String
    Dim TempSplit() As String
    
    TempSplit = Split(Watt, ",")
    
    If UBound(TempSplit) = 19 Then
        If temp > 19 And temp < 151 Then
            CommandBuf = "SETPFCPARAMETER_" & temp & "," & Watt
            TheExec.Datalog.WriteComment Trim(CommandBuf)
            Debug.Print Trim(CommandBuf)
            Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
            TheExec.Datalog.WriteComment Trim(ReplyBuf)
            Debug.Print Trim(ReplyBuf)
        Else
            TheExec.Datalog.WriteComment "Your input temp is out of spec!"
        End If
    Else
        TheExec.Datalog.WriteComment "Your input Watt is not 20 sets!"
    End If
    
End Function

Public Function ATC_Connect() As Long

    If True = ConnectDevice_NI(0, 2, T100s) Then '' 0 is GPIB controler ID, 2 is Handler GPIB Address , T100s is timeout settings.

        TheExec.Datalog.WriteComment "GPIB Connect done!"
        Debug.Print "GPIB Connect done!"
    Else
        TheExec.Datalog.WriteComment "GPIB Connect Fail"
        Debug.Print "GPIB Connect Fail"
    End If

End Function

Public Function Tempread() As Long

Dim CommandBuf As String
Dim ReplyBuf As String

 If True = ConnectDevice_NI(0, 2, T100s) Then '' 0 is GPIB controler ID, 2 is Handler GPIB Address , T100s is timeout settings.

        TheExec.Datalog.WriteComment "GPIB Connect done!"
        Debug.Print "GPIB Connect done!"
    Else
        TheExec.Datalog.WriteComment "GPIB Connect Fail"
        Debug.Print "GPIB Connect Fail"
    End If



CommandBuf = "TEMPARM?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
Debug.Print Trim(ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)

CommandBuf = "SETTEMP?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)


End Function

Public Function TestIPD(InstAdd As Integer, outputfolder As String)
    Dim reply_buf As String
    Dim Resistor_Buf As String
    Dim Z_Buf As String
    Dim C_Buf As String
    Dim WFM As String
    Dim fileNum As Integer
    Dim SplitResult() As String
    Dim CalcResult() As Double
    Dim Vertical_Scale As String
    Dim OffSet_Str As String
    Dim VoltOffSet As Double
    Dim FullScale As Double
    Dim WriteDataToFile As Boolean
    Dim FinalResult As New DSPWave
    Dim FFTResult As New DSPWave
    Dim tmpstr As String
    Dim ymult As Double
    Dim yzero As Double
    Dim yoff  As Double
    Dim xincr As Double
    Dim yscale As Double
    Dim outputfilename As String
    Dim StartPoint As String
    Dim StopPoint As String
    
    Dim i As Integer
    fileNum = freefile()
    WriteDataToFile = True
    Call ConnectDevice_NI(0, InstAdd, T3s)
    
    Call GPIBSendCommand_NI("HEADer off")
    Call GPIBSendCommand_NI("CLS")
    Call GPIBSendQuery_NI("CH3:SCAle?", Vertical_Scale)
''    Call GPIBSendQuery_NI("CH3:OffSet?", OffSet_Str)
    Call GPIBSendCommand_NI("ACQuire:DATA:CLEar")
    Call GPIBSendCommand_NI("ACQuire:STATE ON")
    TheHdw.Wait 1
    ''Get Channel 3 Data
    Call GPIBSendCommand_NI("DATa:SOUrce CH3")
    Call GPIBSendCommand_NI("DATa:ENCdg ascii")
    
    Call GPIBSendCommand_NI("Data:start 1")
    ''Call GPIBSendQuery_NI("Data:start?", startpoint)
    Call GPIBSendCommand_NI("Data:stop 2000")
    ''Call GPIBSendQuery_NI("Data: stop?", stoppoint)
    
    Call GPIBSendQuery_NI("CURVe?", WFM)
    ymult = QueryDoubleResult("wfmoutpre:Ymult?")
    yzero = QueryDoubleResult("wfmoutpre:Yzero?")
    yoff = QueryDoubleResult("wfmoutpre:YOFF?")
    xincr = QueryDoubleResult("wfmoutpre:XINCR?")
    yscale = QueryDoubleResult("wfmoutpre:yscale?")
    Call GPIBSendCommand_NI("ACQuire:STATE OFF")
    Call DisconnectDevice_NI
    
    
   ''WFM = ReadData
   SplitResult = Split(WFM, ",")
   CalculateResult WFM, yzero, yoff, yscale, FinalResult, outputfolder
  
  If (TheExec.enableWord("PlotWaveForm") = True) Then
      FinalResult.SampleRate = 1 / xincr
      FinalResult.Plot
      FFTResult = FinalResult.FFT
      FFTResult.Plot
  End If

End Function
Public Sub PrintOutFormat(CalcResult() As Double, outputfolder As String)
    
    Dim lotId As String
    Dim WaferID As String
    Dim xcoor As String
    Dim ycoor As String
    Dim i As Long
    Dim line As String
    'STEP1. Get lotId/waferId/X/Y Info & output path
    lotId = TheExec.Datalog.Setup.LotSetup.lotId  'ex: A12345
    WaferID = TheExec.Datalog.Setup.WaferSetup.ID  'ex: 12
    xcoor = CStr(TheExec.Datalog.Setup.WaferSetup.GetXCoord(0))  'suppose 1 site only
    ycoor = CStr(TheExec.Datalog.Setup.WaferSetup.GetYCoord(0))
    If lotId = "" Then
      lotId = "A12345"
    End If
    If WaferID = "" Then
      WaferID = "12"
    End If
     If xcoor = "-32768" Then
      xcoor = "1"
    End If
    If ycoor = "-32768" Then
      ycoor = "1"
    End If
    
    Dim outPath As String
    outPath = outputfolder + "\" + lotId + "_" + WaferID + "_" + START_DATE_TIME + ".txt"  'ex: A12345_12_20160805_110512.txt
    'START_DATE_TIME define as a global, and get value at OnProgramStarted()

    'STEP2. Print out
    Dim fileNum As Integer
    fileNum = freefile
    Open outPath For Append Access Write As #fileNum

        line = lotId + "-" + WaferID + "," + xcoor + "," + ycoor
        For i = 0 To UBound(CalcResult) - 1
            line = line + "," + CStr(CalcResult(i))
        Next i
        Print #fileNum, line
    Close #fileNum
    
End Sub

Public Sub CalculateResult(WFM As String, yzero As Double, yoff As Double, yscale As Double, FinalResult As DSPWave, outputfolder As String)

 Dim SpiltResult() As String
 Dim CalcResult() As Double
 Dim WriteDataToFile As Boolean
 Dim i As Long
 Dim tmpstr As String
 Dim outputfilename As String
 Dim lotId As String
 Dim WaferID As String
 Dim DieY As String
 Dim DieX As String
 
 
 
 WriteDataToFile = True
 outputfilename = outputfolder + "\Sample_" + timestamp() + ".txt"
     '''' Write File

 
 
'' If (WFM = "") Then
''  WFM = ReadData("C:\tmp\IPDWaveformCalc_NoHeader_Device.txt")
'' End If
'' tmpStr = Replace(WFM, " ", "")
'' tmpStr = Replace(tmpStr, vbTab, "")
  SpiltResult = Split(WFM, ",")
    ReDim CalcResult(UBound(SpiltResult))
    For i = 0 To UBound(SpiltResult) - 1
       ''tmpStr = CheckNum(SpiltResult(i))
       ''Debug.Print SpiltResult(i) + " " + tmpStr
       CalcResult(i) = ((CLng(SpiltResult(i)) - yoff) * yscale) + yzero
       '''Full Scale 2^16
    Next i
    FinalResult.data = CalcResult
    
    'Print out CalcResult() array in LotId,XCoor,YCoor,RawData... Format
    PrintOutFormat CalcResult, outputfolder
    
    'If WriteDataToFile = True Then
    '       Dim FileNum As Integer
    '       FileNum = FreeFile
    '   ''  Open "c:\tmp\IPDWaveformCalc_NoHeader_Device.txt" For Output Access Write As #FileNum
    '       Open outputfilename For Output Access Write As #FileNum
    '       Print #FileNum, WFM
    '       Print #FileNum, "yoff=" + CStr(yoff) + " yscale=" + CStr(yscale) + " yzero=" + CStr(yzero)
    '       For i = 0 To UBound(SpiltResult) - 1
    '        Print #FileNum, CStr(CalcResult(i))
    '       Next i
    '     Close #FileNum
    'End If

End Sub

Public Function InitInstrument(GPIBAdd As Integer)
    ''' Connect Instrument which GPIB Address is 1, Time out is 1Sec
    Call ConnectDevice_NI(0, GPIBAdd, T3s)
    
    '''' PreSet Channel 3
    Call GPIBSendCommand_NI("TDR:CH3:PRESET")
    '''' Set Channel 3 Meas Volt
    Call GPIBSendCommand_NI("TDR:CH3:UNIts volt")
    '''' Set Vertical Scale = 100 mV/Div
    Call GPIBSendCommand_NI("CH3:SCAle 0.1")
    ''' Set Sample Rate 25k
    Call GPIBSendCommand_NI("TDR:INTRate 25000")
    ''' Set the Open Point at 10% of screen
    Call GPIBSendCommand_NI("HORizontal:Main:REFPoint 0.1")
    ''' Set Open Point at 40us
    Call GPIBSendCommand_NI("HORizontal:Main:Position 40e-6")
    ''' Set Horizontal Scale 5 us/div
    Call GPIBSendCommand_NI("HORizontal:MAIn:SCAle 5e-6")
    '''Call GPIBSendCommand_NI("HORizontal:MAGnify1RECordlength 4000")
    Call DisconnectDevice_NI
End Function
Private Function ReadData(DebugLogFile As String) As String

   Dim nReader As Integer
   Dim LineData As String
   nReader = freefile()

   Open DebugLogFile For Input Access Read As #nReader
    Line Input #nReader, LineData
    Line Input #nReader, LineData
    Close #nReader
    ReadData = LineData
    
End Function
 
 Private Function QueryDoubleResult(GPIBCMD As String) As Double
    Dim tmpstr As String
    Call GPIBSendQuery_NI(GPIBCMD, tmpstr)
    If (tmpstr <> "") Then
       QueryDoubleResult = CDbl(tmpstr)
    Else
       QueryDoubleResult = 0
    End If
 End Function

''Public Function InitialSetup_Old(InstAdd As Integer)
''    Dim reply_buf As String
''    Dim Resistor_Buf As String
''    Dim Z_Buf As String
''    Dim C_Buf As String
''    Dim WFM As String
''    Dim FileNum As Integer
''    Dim SplitResult() As String
''    Dim CalcResult() As Double
''    Dim Vertical_Scale As String
''    Dim OffSet_Str As String
''
''    Dim FullScale As Double
''    Dim i As Integer
''    Dim startpoint As String
''    Dim stoppoint As String
''
''    FileNum = FreeFile()
''
''    Call ConnectDevice_NI(0, InstAdd, T3s)
''    '''PreSet Channel 3
''    Call GPIBSendCommand_NI("TDR:CH3:PRESET")
''    ''' Set Channel 3 Unit to Volt
''    Call GPIBSendCommand_NI("TDR:CH3:UNIts volt")
''    Call GPIBSendCommand_NI("TDR:INTRate 25000")
''    '''' Set Horizontal Scale
''    Call GPIBSendCommand_NI("HORizontal:MAIn:SCAle 5e-6")
''    Call GPIBSendCommand_NI("HORizontal:Main:REFPoint 0.1")
''    Call GPIBSendCommand_NI("HORizontal:Main:Position 40e-6")
''    Call GPIBSendCommand_NI("HEADer off")
''    Call GPIBSendCommand_NI("CH3:SCAle 0.1")
''
''    Call GPIBSendQuery_NI("CH3:SCAle?", Vertical_Scale)
''    Call GPIBSendQuery_NI("CH3:OffSet?", OffSet_Str)
''
''''    Call GPIBSendCommand_NI("HORizontal:MAGnify1RECordlength 1000")
''''    Call GPIBSendCommand_NI("HORizonta2:MAGnify1RECordlength 1000")
''    Call GPIBSendCommand_NI("ACQuire:DATA:CLEar")
''    Call GPIBSendCommand_NI("ACQuire:STATE ON")
''    ''Call GPIBSendCommand_NI("ACQuire:NUMAVg 2")
''
''     Open "c:\tmp\IPDWaveformCalc_NoHeader_Device.txt" For Output Access Write As #FileNum
''           Print #FileNum, "Vertical Scale=" + Vertical_Scale
''    If (Vertical_Scale <> "") Then
''       FullScale = 10 * CDbl(Vertical_Scale)
''    End If
''    ''''' Get Data
''    ''''' Wait Time
''    TheHdw.Wait 1.5
''    Call GPIBSendCommand_NI("DATa:SOUrce CH3")
''    Call GPIBSendCommand_NI("DATa:ENCdg ascii")
''    Call GPIBSendQuery_NI("Data: start 1")
''    Call GPIBSendQuery_NI("Data: start?", startpoint)
''    Call GPIBSendQuery_NI("Data: stop 4000")
''    Call GPIBSendQuery_NI("Data: stop?", stoppoint)
''    Call GPIBSendQuery_NI("CURVe?", WFM)
''
''    '''''
''    SplitResult = Split(WFM, ",")
''  ''  ReDim CalcResult(UBound(SplitResult))
''''    For i = 0 To UBound(SplitResult)
''''       CalcResult(i) = FullScale * CLng(SplitResult(i)) / (2 ^ 16)
''''       '''Full Scale 2^16
''''    Next i
''      '' Open "c:\tmp\IPDWaveformRaw.txt" For Output Access Write As #FileNum
''      '''Print #FileNum, "FullScale=" + CStr(FullScale)
''      Print #FileNum, WFM
''     Close #FileNum
''
''''      Print #FileNum, WFM
''''    For i = 0 To UBound(CalcResult)
''''      Print #FileNum, CStr(CalcResult(i))
''''    Next i
''''    Close #FileNum
''''
''''
''''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:STATE on")
''''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:TYPe Pcross")
''''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:REFLevel1:ABSolute:MID 350mv")
''''''    Call GPIBSendQuery_NI("MEASUrement:MEAS1:VALue?", reply_buf)
''''''    Call GPIBSendCommand_NI("HORizontal: Main: REFPoint " + ReplyBuf)
''''    TheExec.Datalog.WriteComment ("Ref Point=>" + reply_buf)
''''
''''     ''' Meas R
''''    Call GPIBSendCommand_NI("TDR:CH1:UNIts ohm")
''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:STATE on")
''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:TYPe minimum")
''''    Call GPIBSendQuery_NI("MEASUrement:MEAS1:VALue?", Resistor_Buf)
''''    TheExec.Datalog.WriteComment ("R=>" + Resistor_Buf)
''''
''''     ''' Meas Maximum
''''    Call GPIBSendCommand_NI("TDR:CH1:UNIts volt")
''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:STATE on")
''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:TYPe maximum")
''''    Call GPIBSendQuery_NI("MEASUrement:MEAS1:VALue?", Z_Buf)
''''    TheExec.Datalog.WriteComment ("Z=>" + Z_Buf)
''''    ''' Meas C
''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:STATE on")
''''    Call GPIBSendCommand_NI("MEASUrement:MEAS1:TYPe Pcross")
''''    Call GPIBSendCommand_NI("MEASUrement:MEAS<x>:REFLevel<x>:ABSolute:MID " + CStr(0.63 * CDbl(Z_Buf)))
''''    Call GPIBSendQuery_NI("MEASUrement:MEAS1:VALue?", C_Buf)
''''        TheExec.Datalog.WriteComment ("C=>" + C_Buf)
''''''TDR:CH<x>:UNIts volt
''''TDR:INTRate 25k
''''HORizontal:MAIn:SCAle 0.1us
''''MEASUrement:MEAS<x>:STATE on
''''MEASUrement:MEAS<x>:TYPe Pcross
''''MEASUrement:MEAS<x>:REFLevel<x>:ABSolute:MID 350mv
''''MEASUrement:MEAS<x>:VALue? =x
''''HORizontal: Main: REFPoint X
''''After probed
''''TDR:CH<x>:UNIts ohm
''''MEASUrement:MEAS<x>:STATE on
''''MEASUrement:MEAS<x>:TYPe minimum
''''MEASUrement:MEAS<x>:VALue?
''''Above  measure  resistant
''''TDR:CH<x>:UNIts volt
''''MEASUrement:MEAS<x>:STATE on
''''MEASUrement:MEAS<x>:TYPe maximum
''''MEASUrement:MEAS<x>:VALue? =Z
''
''''Check the percentage of 63% of Z =A
''''MEASUrement:MEAS<x>:STATE on
''''MEASUrement:MEAS<x>:TYPe Pcross
''''MEASUrement:MEAS<x>:REFLevel<x>:ABSolute:MID A
''''MEASUrement:MEAS<x>:VALue? =C
''    ''Call GPIBSendQuery_NI("*IDN?", reply_buf)
''    Call DisconnectDevice_NI
''End Function
Public Function CheckNum(data As String) As String
  Dim i As Integer
  Dim strLen As Integer
  Dim tmpstr As String
  Dim tmpChar As String
  tmpstr = ""
  strLen = Len(data)
  For i = 1 To strLen
   tmpChar = mid(data, i, 1)
   If (IsNumeric(tmpChar) = True Or tmpChar = "-") Then
    tmpstr = tmpstr + tmpChar
   End If
  Next i
  CheckNum = tmpstr
  ''Debug.Print tmpStr
End Function

Public Function timestamp() As String
timestamp = CStr(Year(Now())) + CStr(Month(Now())) + CStr(Day(Now())) + "_" + CStr(Hour(Now())) + CStr(Minute(Now())) + CStr(Second(Now()))
End Function

Public Function GPIB_Ctrl_Test_CS(Optional temp As String, Optional targettemp As String) As Long

Dim CommandBuf As String
Dim ReplyBuf As String
Dim StartTime As Double
Dim endTime As Double
Dim runcount As Long


runcount = 0

If True = ConnectDevice_NI(0, 2, T100s) Then '' 0 is GPIB controler ID, 2 is Handler GPIB Address , T100s is timeout settings.

    TheExec.Datalog.WriteComment "GPIB Connect done!"
Else
    TheExec.Datalog.WriteComment "GPIB Connect Fail"
End If

CommandBuf = "SETTEMP?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)

Do

StartTime = Timer()

CommandBuf = "SETTEMP +25"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)


    
                                                       
CommandBuf = "DEVICETEMP_ " & temp
TheExec.Datalog.WriteComment Trim(CommandBuf)
Debug.Print Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)
Debug.Print Trim(ReplyBuf)


endTime = Timer()
TheExec.Datalog.WriteComment "Setup Temp Cost " & CStr(endTime - StartTime) & " sec"

StartTime = Timer()

CommandBuf = "TEMPARM?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)

Dim ReadTemp() As String
ReadTemp() = Split(ReplyBuf, "_")

'If ReadTemp(2) < targettemp Then
'
'goto

runcount = runcount + 1

TheHdw.Wait 60

Loop Until ReadTemp(2) > targettemp Or ReadTemp(3) > targettemp Or runcount = 8




endTime = Timer()
TheExec.Datalog.WriteComment CStr(endTime - StartTime) & " sec"

CommandBuf = "SETTEMP?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)
    
    
End Function

Public Function GPIB_Ctrl_Test_forloop(Optional temp As String) As Long

Dim CommandBuf As String
Dim ReplyBuf As String
Dim StartTime As Double
Dim endTime As Double
Dim runcount As Long
Dim temp_target_ANE2 As New SiteDouble
Dim temp_target_ANE1 As New SiteDouble
Dim Num_target As String
Dim flag_run As Boolean




Dim StepIndex_Val As Long


'For Each site In TheExec.sites
'
'temp_target_ANE2(site) = GetStoredData("ts_ane2_trim_25c")
'temp_target_ANE1(site) = GetStoredData("ts_ane1_trim_25c")
'
'
'
'Next site

temp_target_ANE2(0) = GetStoredData("ts_ane2_trim_25c")

Num_target = temp_target_ANE2(0)

StepIndex_Val = CDbl(val(TheExec.flow.var("SrcCodeIndx").value))

TheExec.Datalog.WriteComment "SrcCodeIndx:   " & StepIndex_Val


runcount = 0

If True = ConnectDevice_NI(0, 2, T100s) Then '' 0 is GPIB controler ID, 2 is Handler GPIB Address , T100s is timeout settings.

    TheExec.Datalog.WriteComment "GPIB Connect done!"
Else
    TheExec.Datalog.WriteComment "GPIB Connect Fail"
End If

CommandBuf = "SETTEMP?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)

'///////////////////////////////////////////////// For safity don't increase temp over limit///////////////////////////
If temp_target_ANE2(0) > 90 Then
    
    CommandBuf = "SETTEMP +25"
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    
    
    Exit Function   'or temp_target_ANE2(0).math.Multiply
    End If
    
 '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    CommandBuf = "TEMPARM?"
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    Dim k As Long
    Dim ReadTemp() As String
    ReadTemp() = Split(ReplyBuf, "_")
    
    flag_run = False
    
    If Num_target - ReadTemp(2) < 20 Then
    
        If Abs(StepIndex_Val - Num_target) > 10 Then

'            For k = 0 To 10
'                Wait 1
'
'
'                If Abs(StepIndex_Val - Num_target) < 10 Then
'                    Exit For
'                Else
                    flag_run = True
'                End If
'            Next k
        End If
    
  If flag_run = False Then
  
   ' StartTime = Timer()
'    StepIndex_Val = 60
     CommandBuf = "SETTEMP +" & StepIndex_Val
'     CommandBuf = "SETTEMP +40"
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    
        
                                                           
    CommandBuf = "DEVICETEMP_ " & temp
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Debug.Print Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    Debug.Print Trim(ReplyBuf)
    
    
    'EndTime = Timer()
    'TheExec.DataLog.WriteComment "Setup Temp Cost " & CStr(EndTime - StartTime) & " sec"
    
    'StartTime = Timer()
    
    CommandBuf = "TEMPARM?"
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    
    'Dim ReadTemp() As String
    'ReadTemp() = Split(ReplyBuf, "_")
    
    'If ReadTemp(2) < targettemp Then
    '
    'goto
    
    
    'thehdw.Wait 60
    
    
    
    
    
    
    'EndTime = Timer()
    'TheExec.DataLog.WriteComment CStr(EndTime - StartTime) & " sec"
    
    CommandBuf = "SETTEMP?"
    TheExec.Datalog.WriteComment Trim(CommandBuf)
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    TheExec.Datalog.WriteComment Trim(ReplyBuf)
    
    TheHdw.Wait 2
        
End If
End If

    
End Function

Public Function GPIB_Ctrl_Test() As Long

Dim CommandBuf As String
Dim ReplyBuf As String
Dim StartTime As Double
Dim endTime As Double

If True = ConnectDevice_NI(0, 2, T100s) Then '' 0 is GPIB controler ID, 2 is Handler GPIB Address , T100s is timeout settings.

    TheExec.Datalog.WriteComment "GPIB Connect done!"
Else
    TheExec.Datalog.WriteComment "GPIB Connect Fail"
End If

CommandBuf = "SETTEMP?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)


StartTime = Timer()

CommandBuf = "SETTEMP +25"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)





endTime = Timer()
TheExec.Datalog.WriteComment "Setup Temp Cost " & CStr(endTime - StartTime) & " sec"

StartTime = Timer()

CommandBuf = "TEMPARM?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)

endTime = Timer()
TheExec.Datalog.WriteComment CStr(endTime - StartTime) & " sec"

CommandBuf = "SETTEMP?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)
    
End Function
Public Function GPIB_Ctrl_Test_CS2() As Long

Dim CommandBuf As String
Dim ReplyBuf As String
Dim StartTime As Double
Dim endTime As Double

If True = ConnectDevice_NI(0, 2, T100s) Then '' 0 is GPIB controler ID, 2 is Handler GPIB Address , T100s is timeout settings.

    TheExec.Datalog.WriteComment "GPIB Connect done!"
Else
    TheExec.Datalog.WriteComment "GPIB Connect Fail"
End If

CommandBuf = "SETTEMP?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)


StartTime = Timer()

CommandBuf = "SETTEMP +45"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)


endTime = Timer()
TheExec.Datalog.WriteComment "Setup Temp Cost " & CStr(endTime - StartTime) & " sec"

CommandBuf = "DEVICETEMP_ " & "10,10"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Debug.Print Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)
Debug.Print Trim(ReplyBuf)


StartTime = Timer()

CommandBuf = "TEMPARM?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)

endTime = Timer()
TheExec.Datalog.WriteComment CStr(endTime - StartTime) & " sec"

CommandBuf = "SETTEMP?"
TheExec.Datalog.WriteComment Trim(CommandBuf)
Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
TheExec.Datalog.WriteComment Trim(ReplyBuf)
    
End Function
Public Sub create_folder(ByVal folder_path As String)
        If Dir(folder_path, vbDirectory) = "" Then
            Debug.Print ("Create " + folder_path)
            MkDir (folder_path)
        Else
            Debug.Print (folder_path & " exists")
        End If
End Sub



Public Function ATC_set_tempoffset(temp As String)
    
    'ATC_ControlTemp ("-3,-3,-3,-3")
    ATC_ControlTemp (temp)
    'ATC_ControlTemp ("-9,-9,-9,-9")
    TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & ";Temp offset"
End Function

Public Function ATC_rollback_oritemp()
    
    ATC_ControlTemp ("0,0,0,0")

End Function


Public Function GPIB_PID_Ctrl(Optional PID_Ctrl As String)

    Dim pid() As String
    
    Dim PID_P_Val As Long
    Dim PID_I_Val As Long
    Dim PID_D_Val As Long
    
'''    If PID_Ctrl = "" Then
'''        PID_P_Val = CDbl(TheExec.flow.var("Coefficient_P").value)
'''        PID_I_Val = CDbl(TheExec.flow.var("Coefficient_I").value)
'''        PID_D_Val = CDbl(TheExec.flow.var("Coefficient_D").value)
'''        PID_Ctrl = PID_P_Val & "," & PID_I_Val & "," & PID_D_Val
'''    End If

    'PID_Ctrl = PID_P_Val & "," & PID_I_Val & "," & PID_D_Val
    PID_Ctrl = "400,0,400"    'ori= "300,5,250"  / "400,0,400" "400,1,400"good
    
    TheExec.Datalog.WriteComment "PID Condition:" & PID_Ctrl
    pid = Split(PID_Ctrl, ",")
    Call GPIBCtrl.Set_PID(pid)
    'TheExec.DataLog.WriteComment "**********************GPIB Offset Setup Done**********************"

End Function


Public Function GPIB_PID_rollback(Optional PID_Ctrl As String)

    Dim pid() As String
    
    Dim PID_P_Val As Long
    Dim PID_I_Val As Long
    Dim PID_D_Val As Long
    
'''    If PID_Ctrl = "" Then
'''        PID_P_Val = CDbl(TheExec.flow.var("Coefficient_P").value)
'''        PID_I_Val = CDbl(TheExec.flow.var("Coefficient_I").value)
'''        PID_D_Val = CDbl(TheExec.flow.var("Coefficient_D").value)
'''        PID_Ctrl = PID_P_Val & "," & PID_I_Val & "," & PID_D_Val
'''    End If

    'PID_Ctrl = PID_P_Val & "," & PID_I_Val & "," & PID_D_Val
    PID_Ctrl = "400,0,400"     'ori= "300,5,250"
        
    TheExec.Datalog.WriteComment "PID Condition:" & PID_Ctrl
    pid = Split(PID_Ctrl, ",")
    Call GPIBCtrl.Set_PID(pid)
    'TheExec.DataLog.WriteComment "**********************GPIB Offset Setup Done**********************"

End Function

Public Function PPMU_Continuity_diode()

    
    Dim PPMUMeasure As New PinListData
    Dim i As Long

'''''     TheHdw.DCVS.Pins("VDD_CIO").Gate = False
''''''
''''''    '    TheHdw.Digital.Pins(digital_pins).Disconnect
''''''    '
''''''    '    If connect_all_pins <> "" Then
'''''            TheHdw.Digital.Pins("ATCPHY3_AUX_N_MEAS").Disconnect
'''''            TheHdw.Digital.Pins("ATCPHY3_AUX_N").Disconnect
'''''            TheHdw.PPMU.Pins("ATCPHY3_AUX_N").Disconnect
'''''''            With TheHdw.PPMU.Pins("ATCPHY3_AUX_N")
'''''''                .ForceV 0#
'''''''                .Connect
'''''''                .Gate = tlOn
'''''''            End With
'''''''    '    End If
''''''        '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
'''''
'''''    With TheHdw.PPMU.Pins("ATCPHY3_AUX_N_MEAS")
'''''        .ForceV 0#
'''''        .Connect
'''''        .Gate = tlOn
'''''    End With
'''''
'''''    TheHdw.PPMU.Pins("ATCPHY3_AUX_N_MEAS").ForceI 100 * uA
'''''    TheHdw.Wait 0.005
'''''    TheHdw.Utility("k74").State = tlUtilBitOff
'''''    TheHdw.Wait 0.005
'''''
     '''For i = 0 To 100
     If TheExec.flow.enableWord("A_Prober") = True Then
        PPMUMeasure = TheHdw.PPMU.pins(Tj_measurement_pin).Read(tlPPMUReadMeasurements, 20)
        Debug.Print "Site0 : " & PPMUMeasure.pins(Tj_measurement_pin).value(0)
        ''TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temp =" & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(0)
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Measurement =" & PPMUMeasure.pins(Tj_measurement_pin).value(0)
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temperature =" & (TJ_Slope_Cal * (PPMUMeasure.pins(Tj_measurement_pin).value(0)) + TJ_Offset_Cal)
     
     Else
        PPMUMeasure = TheHdw.PPMU.pins(Tj_measurement_pin).Read(tlPPMUReadMeasurements, 20)
        Debug.Print "Site1 : " & PPMUMeasure.pins(Tj_measurement_pin).value(1)
        ''TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temp =" & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(0)
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Measurement =" & PPMUMeasure.pins(Tj_measurement_pin).value(1)
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temperature =" & (TJ_Slope_Cal * (PPMUMeasure.pins(Tj_measurement_pin).value(1)) + TJ_Offset_Cal)

     End If
    ''    Debug.Print "Site1 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(1)
    ''    Debug.Print "Site2 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(2)
    ''    Debug.Print "Site3 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(3)
    '''TheHdw.Wait 0.1
        
    '''Next i
'''''    TheHdw.Utility("k74").State = tlUtilBitOn
    '''TheHdw.Wait 0.005
    
'
'    For i = 0 To 9
'        PPMUMeasure = TheHdw.PPMU.Pins("ATCPHY3_AUX_N_MEAS").Read(tlPPMUReadMeasurements, 20)
'        Debug.Print "RelayOn" & "    " & PPMUMeasure.Pins("ATCPHY3_AUX_N_MEAS").Value(0)
'    Next i




End Function

Public Function Pre_Tjmode_setup()

    Dim PPMUMeasure As New PinListData
    Dim i As Long
    
    Dim Temperature_junction As Double
    Dim TJ_CalOffset As Double
    Dim TJ_Slope_Npin As Double:: TJ_Slope_Npin = 1644.3
    Dim TJ_Slope_Ppin As Double:: TJ_Slope_Ppin = 1666.7               '1627.5
    Dim TJ_Offset_Npin As Double:: TJ_Offset_Npin = 419.01
    Dim TJ_Offset_Ppin As Double:: TJ_Offset_Ppin = 395.83              '415.8
    Dim SlopeOffset_cmd As String
    
    Tj_measurement_pin = "ATCPHY3_AUX_P" '''ATCPHY3_AUX_P
    
    TheHdw.DCVS.pins("VDD_CIO_PS").Gate = False
    TheHdw.Digital.pins("ATCPHY3_AUX_N").Disconnect
    TheHdw.Digital.pins("ATCPHY3_AUX_P").Disconnect
   ' TheHdw.PPMU.Pins("ATCPHY3_AUX_P_MEAS").Disconnect
    If Tj_measurement_pin = "ATCPHY3_AUX_N" Then
        TheHdw.Utility("k200").State = tlUtilBitOn   'K201 for P, K200 for N
        TheHdw.Utility("k201").State = tlUtilBitOff   'K201 for P, K200 for N
    ElseIf Tj_measurement_pin = "ATCPHY3_AUX_P" Then
        TheHdw.Utility("k201").State = tlUtilBitOn   'K201 for P, K200 for N
        TheHdw.Utility("k200").State = tlUtilBitOff   'K201 for P, K200 for N
    End If
    TheHdw.Wait 0.005
     

    
    '=================initial=================
     With TheHdw.PPMU.pins(Tj_measurement_pin)
        .ForceV 0#
        .Connect
        .Gate = tlOn
    End With
    
    '=================Force I=================
    TheHdw.PPMU.pins(Tj_measurement_pin).ForceI -200 * uA
    TheHdw.Wait 5
    
    '=================Measure V=================
    'For i = 0 To 500
        
    If TheExec.flow.enableWord("A_Prober") = True Then
        PPMUMeasure = TheHdw.PPMU.pins(Tj_measurement_pin).Read(tlPPMUReadMeasurements, 20)
        Debug.Print "Site0 : " & PPMUMeasure.pins(Tj_measurement_pin).value(0)
        ''TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temp =" & PPMUMeasure.Pins("ATCPHY3_AUX_P").value(0)
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Measurement =" & PPMUMeasure.pins(Tj_measurement_pin).value(0)
    ''    Debug.Print "Site1 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(1)
    ''    Debug.Print "Site2 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(2)
    ''    Debug.Print "Site3 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(3)
        TheHdw.Wait 0.1
    Else
        PPMUMeasure = TheHdw.PPMU.pins(Tj_measurement_pin).Read(tlPPMUReadMeasurements, 20)
        Debug.Print "Site1 : " & PPMUMeasure.pins(Tj_measurement_pin).value(1)
        ''TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temp =" & PPMUMeasure.Pins("ATCPHY3_AUX_P").value(0)
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Measurement =" & PPMUMeasure.pins(Tj_measurement_pin).value(1)
    ''    Debug.Print "Site1 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(1)
    ''    Debug.Print "Site2 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(2)
    ''    Debug.Print "Site3 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(3)
        TheHdw.Wait 0.1
    End If
   ' Next i
    TheHdw.PPMU.AllowPPMUFuncRelayConnection True, False

    
    '=============================================================Calbration
    If Tj_measurement_pin = "ATCPHY3_AUX_N" Then
        Temperature_junction = Round(TJ_Slope_Npin * PPMUMeasure.pins(Tj_measurement_pin).value(0) + TJ_Offset_Npin, 2)
        TJ_CalOffset = Temperature_junction - Tj_Tester
        
        TJ_Offset_Cal = Round(TJ_Offset_Npin - TJ_CalOffset, 2)
        TJ_Slope_Cal = TJ_Slope_Npin
        
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temperature =" & (TJ_Slope_Cal * (PPMUMeasure.pins(Tj_measurement_pin).value(0)) + TJ_Offset_Cal)
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Slope_Cal =" & TJ_Slope_Cal
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Offset_Cal =" & TJ_Offset_Cal
    ElseIf Tj_measurement_pin = "ATCPHY3_AUX_P" Then
        Temperature_junction = Round(TJ_Slope_Ppin * PPMUMeasure.pins(Tj_measurement_pin).value(0) + TJ_Offset_Ppin, 2)
        TJ_CalOffset = Temperature_junction - Tj_Tester
        
        TJ_Offset_Cal = Round(TJ_Offset_Ppin - TJ_CalOffset, 2)
        TJ_Slope_Cal = TJ_Slope_Ppin
        
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temperature =" & (TJ_Slope_Cal * (PPMUMeasure.pins(Tj_measurement_pin).value(0)) + TJ_Offset_Cal)
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Slope_Cal =" & TJ_Slope_Cal
        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Offset_Cal =" & TJ_Offset_Cal
    End If

''    TheHdw.PPMU.Pins("ATCPHY3_AUX_N").ForceI -150 * uA
''    TheHdw.Wait 0.005
''
''    '=================Measure V=================
''    For i = 0 To 500
''        PPMUMeasure = TheHdw.PPMU.Pins("ATCPHY3_AUX_N").Read(tlPPMUReadMeasurements, 20)
''        Debug.Print "Site0 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(0)
''        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temp =" & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(0)
''    ''    Debug.Print "Site1 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(1)
''    ''    Debug.Print "Site2 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(2)
''    ''    Debug.Print "Site3 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(3)
''        TheHdw.Wait 0.1
''    Next i
''
''        TheHdw.PPMU.Pins("ATCPHY3_AUX_N").ForceI -100 * uA
''    TheHdw.Wait 0.005
''
''    '=================Measure V=================
''    For i = 0 To 500
''        PPMUMeasure = TheHdw.PPMU.Pins("ATCPHY3_AUX_N").Read(tlPPMUReadMeasurements, 20)
''        Debug.Print "Site0 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(0)
''        TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temp =" & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(0)
''    ''    Debug.Print "Site1 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(1)
''    ''    Debug.Print "Site2 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(2)
''    ''    Debug.Print "Site3 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(3)
''        TheHdw.Wait 0.1
''    Next i
    '========================================================ATCPHY3_AUX_N================
'''''''''''''''''    thehdw.PPMU.Pins("ATCPHY3_AUX_P").Disconnect
'''''''''''''''''
'''''''''''''''''    thehdw.DCVS.Pins("VDD_CIO").Gate = False
'''''''''''''''''    thehdw.Digital.Pins("ATCPHY3_AUX_N_MEAS").Disconnect
'''''''''''''''''    thehdw.Digital.Pins("ATCPHY3_AUX_N").Disconnect
'''''''''''''''''    thehdw.PPMU.Pins("ATCPHY3_AUX_N_MEAS").Disconnect
'''''''''''''''''    'thehdw.Utility("k74").State = tlUtilBitOn
''''''''''''''''''''''    TheHdw.Utility("k74").State = tlUtilBitOff
'''''''''''''''''    thehdw.Wait 0.005
'''''''''''''''''
'''''''''''''''''     With thehdw.PPMU.Pins("ATCPHY3_AUX_N")
'''''''''''''''''        .ForceV 0#
'''''''''''''''''        .Connect
'''''''''''''''''        .Gate = tlOn
'''''''''''''''''    End With
'''''''''''''''''
'''''''''''''''''    thehdw.PPMU.Pins("ATCPHY3_AUX_N").ForceI 200 * uA
'''''''''''''''''    thehdw.Wait 0.005
'''''''''''''''''
'''''''''''''''''    For i = 0 To 9
'''''''''''''''''        PPMUMeasure = thehdw.PPMU.Pins("ATCPHY3_AUX_N").Read(tlPPMUReadMeasurements, 20)
'''''''''''''''''        Debug.Print "Site0 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(0)
'''''''''''''''''    ''    Debug.Print "Site1 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(1)
'''''''''''''''''    ''    Debug.Print "Site2 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(2)
'''''''''''''''''    ''    Debug.Print "Site3 : " & PPMUMeasure.Pins("ATCPHY3_AUX_N").value(3)
'''''''''''''''''    Next i
    
    '========================================================ATCPHY3_AUX_N================
    
End Function

Public Function Set_ControlMode_VI(OnOff As Boolean)
    Dim Cmd As String:: Cmd = "POWERFOLLOWING"
    Dim parameter As String:: parameter = " _0"
    Dim reply As String
    Dim result As Boolean:: result = False
    
    If OnOff = True Then parameter = " _1"
    Cmd = Cmd & parameter
    
    ' When the GPIB is not conneted, then reconnect
    'If connectted_flag = False Then Call GPIB_Connect
    
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
    
    If right(Cmd, 1) = "1" Then
        TheExec.Datalog.WriteComment " VI mode = " & "True"
    Else
        TheExec.Datalog.WriteComment " VI mode = " & "False"
    End If
End Function

Public Function thermal_control_time()

If TheExec.flow.enableWord("A_PreTrigger") = True Then
    If TheExec.DataManager.instancename = "EVS_Static_Power_Ramp_Multi_GFX_SCAN_01_1p25_TJ" = True Then
        
        ATC_ControlTemp ("-3,-3,-3,-3")
    
    ElseIf TheExec.DataManager.instancename = "EVS_Static_Power_Ramp_Multi_GFX_SCAN_01_1p3_TJ" = True Then
    
    ElseIf TheExec.DataManager.instancename = "EVS_Static_Power_Ramp_Multi_GFX_SCAN_01_1p35_TJ" = True Then
    
    ElseIf TheExec.DataManager.instancename = "EVS_Static_Power_Ramp_Multi_GFX_SCAN_01_1p4_TJ" = True Then
    
    ElseIf TheExec.DataManager.instancename = "EVS_Static_Power_Ramp_Multi_GFX_SCAN_01_1p45_TJ" = True Then
        ATC_ControlTemp ("-6,-6,-6,-6")
    ElseIf TheExec.DataManager.instancename = "EVS_Static_Power_Ramp_Multi_GFX_SCAN_01_1p5_TJ" = True Then
        ATC_ControlTemp ("-8,-8,-8,-8")
    End If
ElseIf TheExec.flow.enableWord("A_PreTrigger") = False Then
    TheHdw.Wait 10
End If

End Function


Public Function Set_VI_table()
    Call ATC_SetPowerTableParameter("80", "100,100,100,100,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")
    Call ATC_SetPowerTableParameter("90", "100,100,100,100,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")
End Function

Public Function Rollback_VI_table()
    Call ATC_SetPowerTableParameter("80", "100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100")
    Call ATC_SetPowerTableParameter("90", "100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100")
End Function

Public Function Fake_VIsignal()

With TheHdw.DCVS.pins("ATC_VREF")
    .Disconnect tlDCVSConnectDefault
    .mode = tlDCVSModeVoltage
    .Voltage.value = 1    ''''4.3
    .Connect tlDCVSConnectDefault
    .Gate = True
End With
TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & ";" & "Fake_signal"

    
End Function
Public Function ATCS_set_tempoffset(temp As String)
    Call GPIBCtrl.HTF_ThermalControl_OffsetTemperature(temp)
    ''TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & ";Temp offset"
End Function


Public Function ATCS_TCR()

If GPIBCtrl.TCR = False Then Call MsgBox("Thermal Control Run(TCR) failed.", vbCritical)

End Function

Public Function ATCS_TJCAL()

If GPIBCtrl.TJCal_HDPlus = False Then Call MsgBox("TJ Caribration failed.", vbCritical)

End Function
Public Function ATCS_TCS()

If GPIBCtrl.TCS = False Then Call MsgBox("Thermal Control Run(TCS) failed.", vbCritical)

End Function

Public Function Printout_ATC_info(ATC_info As String)

    TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & ";" & ATC_info
End Function
Public Function Random_power_presetup()

    'KKK = 0
End Function
Public Function ATCS_TJVOL()

'If GPIBCtrl.TJVOL = False Then Call MsgBox("Thermal Control Run(TJVOL) failed.", vbCritical)

Call GPIBCtrl.TJVOL
''''''''''''''''''''''HS 220511
TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJVOL_Measurement of ATCS prober =" & TJVOL_measurement
TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJVOL_Temperature of ATCS prober =" & ((TJ_Slope_Cal * TJVOL_measurement) + TJ_Offset_Cal)


Dim PPMUMeasure As New PinListData
PPMUMeasure = TheHdw.PPMU.pins(Tj_measurement_pin).Read(tlPPMUReadMeasurements, 20)
TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Measurement of Device  =" & PPMUMeasure.pins(Tj_measurement_pin).value(0)
TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Temperature of Device =" & (TJ_Slope_Cal * (PPMUMeasure.pins(Tj_measurement_pin).value(0)) + TJ_Offset_Cal)
End Function
Public Function ATCS_TJCH(ChannelCnt As String, ChannelList As String)

Call GPIBCtrl.TJCH(ChannelCnt, ChannelList)

End Function

Public Function ATCS_TJSLOPE(ChannelNum As String, TemperatureSlope As String, TemperatureOffset As String)

TemperatureSlope = CStr(TJ_Slope_Cal)
TemperatureOffset = CStr(TJ_Offset_Cal)
Call GPIBCtrl.TJSLOPE(ChannelNum, TemperatureSlope, TemperatureOffset)


'==============calibration=====================
If Abs(ATC_TJ_measurement - Tj_Tester) < 1 Then
    Dim TJ_CalOffset As Double
    TJ_CalOffset = ATC_TJ_measurement - Tj_Tester
    
    TJ_Offset_Cal = Round(TJ_Offset_Cal - TJ_CalOffset, 2)
    TJ_Slope_Cal = TemperatureSlope

    TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Slope_Cal =" & TJ_Slope_Cal
    TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; TJ_Offset_Cal =" & TJ_Offset_Cal
    TheExec.Datalog.WriteComment "           Current time is : " & Strings.Format(Time, "hh:mm:ss") & "." & Strings.right(Strings.Format(Timer, "#0.00"), 2) & "; cal again"

    TemperatureSlope = CStr(TJ_Slope_Cal)
    TemperatureOffset = CStr(TJ_Offset_Cal)
    Call GPIBCtrl.TJSLOPE(ChannelNum, TemperatureSlope, TemperatureOffset)
End If
'==============calibration=====================

End Function
