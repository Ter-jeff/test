Attribute VB_Name = "VBT_PC_Diagnosis"
Option Explicit
'====modify for change path Hi limt  20190807====
Type PinMeasData
   PinName As String
   Threshold As New SiteDouble
   MeasVal As New SiteDouble
   MeasPF As New SiteLong
   FinishSearch As New SiteBoolean
End Type

Type ExecData
   CurrMeasData() As PinMeasData
   PrevUpperData() As PinMeasData
   PrevLowerData() As PinMeasData
End Type

Type LimitSet
   PinName As String
   hiLimit As String
   loLimit As String
   unit As String
End Type
'Public Site As Variant
Public Dict_Res_Hlimit As New Dictionary
Public Parsing_Table_flag As Long
Public m_sAllDigitalPinList As String
Public Relay_On_Pin_Dic_open As New Dictionary ''for digital open test, store need relay on pin. adding 191121
Public Relay_On_Pin_Dic_short As New Dictionary ''for digital open test, store need relay on pin. adding 191121
Public Relay_On_Pin_Dic_short_UVI80 As New Dictionary ''for digital open test, store need relay on pin. adding 191121

Private Dic_Limit As New Dictionary
Private ImpLimit() As LimitSet
Private PrintDebug As Integer
Private Const Zin = 50
Private Const Vih = 2
Private Const ExpectedDriveEdge = 50


Public Function Parsing_Limit_Table() As Long
   '---------------------------------parseing resistance limit table 20190807-------
    Dim wb_def As Worksheet
    Dim wb As Workbook
    Dim Res_Row As Long, Res_Col As Long
    'dim Site
    Dim Pin_KeyName As String
    Dim Pin_HLimit As Double

    Res_Row = 2
    Parsing_Table_flag = 0 '20200408 Debug
   If (Parsing_Table_flag = 0) Then
        Set wb = Application.ActiveWorkbook
        Set wb_def = wb.Sheets("Path_Resistance_limit")
        wb_def.Select
        While wb_def.Cells(Res_Row, 1).value <> "End"
             Pin_KeyName = CStr(wb_def.Cells(Res_Row, 1).value) & UCase(wb_def.Cells(Res_Row, 2).value)
             Pin_HLimit = CDbl(wb_def.Cells(Res_Row, 3).value)
             If Dict_Res_Hlimit.Exists(Pin_KeyName) Then
                Dict_Res_Hlimit.Remove (Pin_KeyName)
             End If
              Dict_Res_Hlimit.Add Pin_KeyName, Pin_HLimit
              Res_Row = Res_Row + 1
        Wend
        Parsing_Table_flag = 1
    End If
   '---------------------------------------------------------------------
End Function

Public Function Trace_Resistance_Meas(TestPin As String, RelayOn_Pin As String) As Long

Dim MeasureValue As New PinListData
Dim ForceValI As Double
Dim MeasurePinStr As String

Dim PinArr() As String
Dim PinCount As Long
Dim pin  As Variant
Dim Res_Value As Double
Dim i As Integer
Dim hiLimit As Double
Dim LowLimit As Double
Dim force_i As Double
Dim Tname As String
Dim sPin As String
On Error GoTo errHandler
hiLimit = 10#
LowLimit = 0

MeasurePinStr = TestPin
Tname = "Digital Path Resistance"
''''' thehdw.DCVS.Pins("All_power").Connect tlDCVSConnectDefault
'''''                     thehdw.Wait 0.01
'''''                     With thehdw.DCVS.Pins("All_power")
'''''                        .mode = tlDCVSModeVoltage
'''''                        .Voltage.Main.Value = 0#
'''''                        .Voltage.Alt.Value = 0#
'''''                        .Voltage.output = tlDCVSVoltageMain
'''''                        .CurrentRange.Value = 0.02
'''''                        .Meter.mode = tlDCVSMeterCurrent
'''''                        .Meter.CurrentRange = 0.02
'''''                        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
'''''                        .Gate = True
'''''                        thehdw.Wait 1#
'''''                      End With
Dim Relay_on_pin_temp1() As String 'add 191121
Dim Relay_on_pin_temp2() As String 'add 191121
Dim Relay_Name As String

Relay_on_pin_temp1 = Split(RelayOn_Pin, ";") 'add 191121
For i = 0 To UBound(Relay_on_pin_temp1)
    Relay_on_pin_temp2 = Split(Relay_on_pin_temp1(i), "+")
    Relay_On_Pin_Dic_short(Relay_on_pin_temp2(0)) = Relay_on_pin_temp2(1)
Next i
Dim Current_Range As Double
TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
For i = 0 To UBound(PinArr)
'For Each pin In PinArr
    MeasureValue.AddPin (PinArr(i))
    For Each site In TheExec.sites.Selected
             'theexec.Sites(site).selectd
           sPin = PinArr(i)
           
         If Relay_On_Pin_Dic_short.Exists(PinArr(i)) Then
                Relay_Name = Relay_On_Pin_Dic_short(PinArr(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOn
                TheHdw.Wait 0.0001
          
          End If
'''''          If sPin = "DDR0_RREF" Then
'''''             thehdw.Utility.pins("K31").State = tlUtilBitOn
'''''             thehdw.Wait 0.0001
'''''          ElseIf sPin = "DDR1_RREF" Then
'''''              thehdw.Utility.pins("K32").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "DDR2_RREF" Then
'''''              thehdw.Utility.pins("K33").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "DDR3_RREF" Then
'''''              thehdw.Utility.pins("K34").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "MIPI0C_REXT" Then
'''''              thehdw.Utility.pins("K41").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "MIPID_REXT" Then
'''''              thehdw.Utility.pins("K42").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "PCIE_RCAL_P" Then
'''''              thehdw.Utility.pins("K71").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "PCIE_RCAL_N" Then
'''''              thehdw.Utility.pins("K73").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDP_RCAL_P" Then
'''''              thehdw.Utility.pins("K51").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDP_RCAL_N" Then
'''''              thehdw.Utility.pins("K53").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDPRX0_RCAL_P" Then
'''''              thehdw.Utility.pins("K61").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDPRX0_RCAL_N" Then
'''''              thehdw.Utility.pins("K63").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDPRX1_RCAL_P" Then
'''''              thehdw.Utility.pins("K65").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDPRX1_RCAL_N" Then
'''''              thehdw.Utility.pins("K67").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "XI0" Or sPin = "XO0" Then
'''''              thehdw.Utility.pins("K05").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "RT_CLK32768" Then
'''''              thehdw.Utility.pins("K06").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "USB_REXT" Then
'''''              thehdw.Utility.pins("K81").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "DBG_USB_REXT" Then
'''''              thehdw.Utility.pins("K82").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "PAD_MTR_RREF_P" Then
'''''              thehdw.Utility.pins("K91").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "PAD_MTR_RREF_N" Then
'''''              thehdw.Utility.pins("K93").State = tlUtilBitOn
'''''              thehdw.Wait 0.0001
'''''          End If
          
      TheHdw.Digital.pins(PinArr(i)).Disconnect
      With TheHdw.PPMU(PinArr(i))
    
          .Connect
        
          .ForceV 0.1, 0.04
          .Gate = tlOn
           TheHdw.Wait 0.01
           'DebugPrintFunc_PPMU leakage_pins.Value
           MeasureValue = .Read(tlPPMUReadMeasurements)
          ' current_range = thehdw.PPMU.Pins(PinArr(i)).MeasureCurrentRange
            Res_Value = TheHdw.PPMU.pins(PinArr(i)).Voltage.value / MeasureValue.pins(PinArr(i)).value
          .ForceV (0)
          .Gate = tlOff
          .Disconnect
      End With
         '---change judge hi/li limit here for different spec 20190807----
          Dim sPin_Key As String
          Dim bb As Integer
         ' Dim Site As Variant
         ' Dim MeasureValue_site() As New SiteDouble
          
          'MeasureValue_site = pldMeasureValue
         ' For Each Site In theexec.sites.Existing
         If PinArr(i) = "VDD_CPU_Monitor" Then
            PinArr(i) = UCase("VDD_CPU_Monitor")
         End If
        

                     sPin_Key = CStr(site) & CStr(PinArr(i))
                     hiLimit = Dict_Res_Hlimit(sPin_Key)
                     
                   ' theexec.Flow.TestLimit pldMeasureValue, Tname:=strTName, lowVal:=0, hiVal:=dHiLimit, Unit:=8, customUnit:="Ohm"
                     TheExec.flow.TestLimit resultVal:=Res_Value, PinName:=PinArr(i), lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
                     'glb_TestInstance = theexec.DataManager.instancename
                     'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
         ' Next Site
        '---change judge hi/li limit here for different spec 20190807----
             
       
            TheHdw.Wait 0.001
            'MeasureValue.Pins(PinArr(i)) = thehdw.PPMU.Pins(pin).Read(tlPPMUReadMeasurements, 10)
          
           ' thehdw.PPMU.Pins(PinArr(i)).Disconnect

            
            With TheHdw.PPMU.pins(PinArr(i))
  
              .Gate = tlOff
              .Disconnect
            End With
            
         If Relay_On_Pin_Dic_short.Exists(PinArr(i)) Then
                Relay_Name = Relay_On_Pin_Dic_short(PinArr(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOff
                TheHdw.Wait 0.0001
          
          End If
            
'''''          If sPin = "DDR0_RREF" Then
'''''             thehdw.Utility.pins("k31").State = tlUtilBitOff
'''''             thehdw.Wait 0.0001
'''''          ElseIf sPin = "DDR1_RREF" Then
'''''              thehdw.Utility.pins("k32").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "DDR2_RREF" Then
'''''              thehdw.Utility.pins("k33").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "DDR3_RREF" Then
'''''              thehdw.Utility.pins("k34").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "MIPI0C_REXT" Then
'''''              thehdw.Utility.pins("k41").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "MIPID_REXT" Then
'''''              thehdw.Utility.pins("k42").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "PCIE_RCAL_P" Then
'''''              thehdw.Utility.pins("k71").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "PCIE_RCAL_N" Then
'''''              thehdw.Utility.pins("k73").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDP_RCAL_P" Then
'''''              thehdw.Utility.pins("k51").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDP_RCAL_N" Then
'''''              thehdw.Utility.pins("k53").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDPRX0_RCAL_P" Then
'''''              thehdw.Utility.pins("K61").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDPRX0_RCAL_N" Then
'''''              thehdw.Utility.pins("K63").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDPRX1_RCAL_P" Then
'''''              thehdw.Utility.pins("K65").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "LPDPRX1_RCAL_N" Then
'''''              thehdw.Utility.pins("K67").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "XI0" Or sPin = "XO0" Then
'''''              thehdw.Utility.pins("K05").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "RT_CLK32768" Then
'''''              thehdw.Utility.pins("K06").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "USB_REXT" Then
'''''              thehdw.Utility.pins("K81").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "DBG_USB_REXT" Then
'''''              thehdw.Utility.pins("K82").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "PAD_MTR_RREF_P" Then
'''''              thehdw.Utility.pins("K91").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          ElseIf sPin = "PAD_MTR_RREF_N" Then
'''''              thehdw.Utility.pins("K93").State = tlUtilBitOff
'''''              thehdw.Wait 0.0001
'''''          End If
           
     Next site
 'Next pin
  Next i
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Trace_Resistance_Meas"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Digital_Open_Voltage_Meas(TestPin As String, RelayOn_Pin As String, LowLimit As Double, hiLimit As Double) As Long

Dim MeasureValue As New PinListData
Dim pldMeasureValue As New PinListData

Dim pldOpenValue As New PinListData
Dim ForceValI As Double
Dim MeasurePinStr As String

Dim PinArr() As String
Dim PinCount As Long
Dim pin  As Variant
Dim Res_Value As Double
Dim i As Integer
'Dim HiLimit As Double
'Dim LowLimit As Double
Dim force_i As Double
Dim Tname As String
Dim Relay_on_pin_temp1() As String 'add 191121
Dim Relay_on_pin_temp2() As String 'add 191121
Dim Relay_Name As String
On Error GoTo errHandler
'HiLimit = 1.2
'LowLimit = 0.8
Dim sPin As String
MeasurePinStr = TestPin
Tname = "Digital_Open"


Dim Current_Range As Double
Relay_on_pin_temp1 = Split(RelayOn_Pin, ";")  'add 191121
For i = 0 To UBound(Relay_on_pin_temp1)
    Relay_on_pin_temp2 = Split(Relay_on_pin_temp1(i), "+")
    Relay_On_Pin_Dic_open(Relay_on_pin_temp2(0)) = Relay_on_pin_temp2(1)
Next i
TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
For i = 0 To UBound(PinArr)
sPin = PinArr(i)
'For Each pin In PinArr
 ' If UCase(PinArr(i)) = "REFCLK_XI0" Or UCase(PinArr(i)) = "REFCLK_RT_CLK32768" Or UCase(PinArr(i)) = "FRC_BACKUP_1" Or UCase(PinArr(i)) = "FRC_BACKUP_4" _
            Or UCase(PinArr(i)) = "FRC_BACKUP_5" Or UCase(PinArr(i)) = "FRC_BACKUP_2" Or UCase(PinArr(i)) = "REFCLK_MONITORCHANNEL" Or UCase(PinArr(i)) = "FRC_BACKUP_3" Then
            'these pin connected to clock buffer
  'ElseIf UCase(PinArr(i)) = "VDD_PCPU_MONITOR" Or UCase(PinArr(i)) = "VDD_GPU_MONITOR" Or UCase(PinArr(i)) = "VDD_SOC_MONITOR" Or UCase(PinArr(i)) = "VDD_CPU_Monitor" Then
           'short with DCVI/DCVS pins, connect capacitors to gnd. Not test current leakage.
 ' Else 'Remove If 20200410
           MeasureValue.AddPin (PinArr(i))
          
          If Relay_On_Pin_Dic_open.Exists(PinArr(i)) Then
                Relay_Name = Relay_On_Pin_Dic_open(PinArr(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOn
                TheHdw.Wait 0.0001
          
        End If
          
          
'''''          If sPin = "PCIE_RCAL_P" Then                             'mark 191121
'''''                thehdw.Utility.pins("K71").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "PCIE_RCAL_N" Then
'''''                thehdw.Utility.pins("K73").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "LPDP_RCAL_P" Then
'''''                thehdw.Utility.pins("K51").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "LPDP_RCAL_N" Then
'''''                thehdw.Utility.pins("K53").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "LPDPRX0_RCAL_P" Then
'''''                thehdw.Utility.pins("K61").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "LPDPRX0_RCAL_N" Then
'''''                thehdw.Utility.pins("K63").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "LPDPRX1_RCAL_P" Then
'''''                thehdw.Utility.pins("K65").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "LPDPRX1_RCAL_N" Then
'''''                thehdw.Utility.pins("K67").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "XI0" Or sPin = "XO0" Then
'''''                thehdw.Utility.pins("K05").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "RT_CLK32768" Then
'''''                thehdw.Utility.pins("K06").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "PAD_MTR_RREF_P" Then
'''''                thehdw.Utility.pins("K91").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            ElseIf sPin = "PAD_MTR_RREF_N" Then
'''''                thehdw.Utility.pins("K93").State = tlUtilBitOn
'''''                thehdw.Wait 0.0001
'''''            End If
       
      TheHdw.Digital.pins(PinArr(i)).Disconnect

'-------------------------------------------------
      With TheHdw.PPMU(PinArr(i))
           .Connect
          TheHdw.Wait 0.002
          '.SetClampsVHi (1)
          .ClampVHi = 1#
          .ForceI -0.0001, 0.0015
          .Gate = tlOn

          TheHdw.Wait 0.004

      '    .ForceV 0.1, 0.04

      End With
'-------------------------------------------------
          'If UCase(PinArr(i)) = "VDD_CPU_MONITOR" Then

          'Else
      
                pldMeasureValue = TheHdw.PPMU.pins(PinArr(i)).Read(tlPPMUReadMeasurements)
                pldOpenValue.AddPin (PinArr(i))
                pldOpenValue.pins(PinArr(i)) = pldMeasureValue.pins(PinArr(i))
          'End If
            With TheHdw.PPMU.pins(PinArr(i))
  
              .Gate = tlOff
              .Disconnect
            End With
            
            
            
          If Relay_On_Pin_Dic_open.Exists(PinArr(i)) Then
                Relay_Name = Relay_On_Pin_Dic_open(PinArr(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOff
                TheHdw.Wait 0.0001
          
          End If
            
''''''         If sPin = "PCIE_RCAL_P" Then                                'mask 191121
''''''                thehdw.Utility.pins("K71").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "PCIE_RCAL_N" Then
''''''                thehdw.Utility.pins("K73").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "LPDP_RCAL_P" Then
''''''                thehdw.Utility.pins("K51").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "LPDP_RCAL_N" Then
''''''                thehdw.Utility.pins("K53").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "LPDPRX0_RCAL_P" Then
''''''                thehdw.Utility.pins("K61").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "LPDPRX0_RCAL_N" Then
''''''                thehdw.Utility.pins("K63").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "LPDPRX1_RCAL_P" Then
''''''                thehdw.Utility.pins("K65").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "LPDPRX1_RCAL_N" Then
''''''                thehdw.Utility.pins("K67").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "XI0" Or sPin = "XO0" Then
''''''                thehdw.Utility.pins("K05").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "RT_CLK32768" Then
''''''                thehdw.Utility.pins("K06").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "PAD_MTR_RREF_P" Then
''''''                thehdw.Utility.pins("K91").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            ElseIf sPin = "PAD_MTR_RREF_N" Then
''''''                thehdw.Utility.pins("K93").State = tlUtilBitOff
''''''                thehdw.Wait 0.0001
''''''            End If
        'End If
  ' End If 'Remove If 20200410

  Next i
  
  TheExec.flow.TestLimit resultVal:=pldOpenValue, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, customUnit:="V", ForceResults:=tlForceNone
  'glb_TestInstance = theexec.DataManager.instancename
  'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Trace_Resistance_Meas"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Power_Trace_Resistance_Meas(TestPin As PinList, Force_volt As Double) As Long
 Dim PowerPins As String
 Dim p As Variant, Pin_Ary() As String, p_cnt As Long
 Dim Res_Value As Double
 Dim Power_data As New PinListData
 Dim First_meas As New PinListData
 Dim Power_volt As New PinListData
 Dim Power_data_value As Double
 Dim i As Integer
 Dim hiLimit As Double
 Dim LowLimit As Double
 Dim Tname As String
 Dim Max_range As Double
 Dim Current_DSPW As New DSPWave
'-------
 Dim chanString As String
 Dim slotstr() As String
 Dim slot As Long
 Dim Instrument_Type As String
 Dim New_I_Range As Double
 Dim New_I_Range_Fold As Double
 Dim j As Integer
'-------
  Dim DCVSMeasRangeList() As Double
  Dim First_measure As Double
  Dim First_measure_max As Double
 On Error GoTo errHandler
 hiLimit = 10#
 LowLimit = 0
 Tname = "DCVS path"
 
 TheExec.DataManager.DecomposePinList TestPin, Pin_Ary, p_cnt
 
For i = 0 To UBound(Pin_Ary)

   'For Each site In TheExec.sites.Existing
       '--------------------------------init Current Range -----------------------------------------
         Call TheExec.DataManager.GetChannelStringFromPinAndSite(Pin_Ary(i), site, chanString)
         
        
         If chanString = "" Then
             TheExec.Datalog.WriteComment ("Warnning : Please check pin type of  " & Pin_Ary(i) & " in channel map")
         Else
             slotstr = Split(chanString, ".")
             slot = CLng(slotstr(0))
             Instrument_Type = TheHdw.config.Slots(slot).type
         End If
         
         If LCase(Instrument_Type) = "hexvs" Then
             New_I_Range = 15#                    'HexVs use 15A
             New_I_Range_Fold = 5#
             
         Else
           New_I_Range = TheHdw.DCVS.pins(Pin_Ary(i)).CurrentRange.Max 'UVS256 use max range
           New_I_Range_Fold = New_I_Range '* 0.1
         End If
       '----------------------------------------------------------------------------------------------
   ' For Each site In theexec.sites.Existing
   
'====================Current profile setup to check current stable time =================
'''''         TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.Signals.Add "DCVSCaptureLabel"
'''''         TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.Signals.DefaultSignal = "DCVSCaptureLabel"
'''''        With TheHdw.DCVS.Pins(Pin_Ary(i))
'''''            .Connect tlDCVSConnectDefault
'''''            .Gate = True
'''''        End With
'''''         With TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.Signals.Item("DCVSCaptureLabel")
'''''
'''''           .Reinitialize
'''''           .Range = 5.6
'''''           .mode = tlDCVSMeterCurrent
'''''           .mode.mode = tlSignalModeUseLoadedValue
'''''           .SampleRate = 200000 'SampleRate
'''''           .SampleSize = 16000 'SampleSize
'''''           .LoadSettings
'''''           .Trigger
'''''         End With
'==========================================================================================

           With TheHdw.DCVS.pins(Pin_Ary(i))
                    .Connect
                    .mode = tlDCVSModeVoltage
                    .CurrentRange.value = New_I_Range ' 0.7 ' 0.0024
                    .CurrentLimit.Source.FoldLimit.level.value = New_I_Range_Fold '1 * .CurrentLimit.Source.FoldLimit.Level.max
                    .CurrentLimit.Sink.FoldLimit.level.value = 1 * .CurrentLimit.Sink.FoldLimit.level.Max
                    .Voltage.Main.value = Force_volt '0.002
                    .Voltage.Alt.value = Force_volt  '0.002
                    .Voltage.Output = tlDCVSVoltageMain
                    .CurrentRange.value = New_I_Range
                    .Meter.mode = tlDCVSMeterCurrent
                    .Meter.CurrentRange = New_I_Range
                    .Meter.Filter.bypass = False
                    .Meter.VoltageRange = 0.05
                    
                    .Gate = True
                    TheHdw.Wait 0.02
            End With
             
               First_meas = TheHdw.DCVS.pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage) 'First measure I for adjust new current range
               
               '''''Print out firt measure current
               'theexec.Datalog.WriteComment ("First Current Range : " & New_I_Range & "A")
               'theexec.Datalog.WriteComment ("First current : ") '& First_meas & "A")
               'theexec.Flow.TestLimit resultVal:=First_meas, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitCustom, ScaleType:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=force_v, ForceUnit:=unitVolt, customUnit:="A", ForceResults:=tlForceNone


            TheHdw.DCVS.pins(Pin_Ary(i)).Gate = False
            '======================Check first measure current then change  current range=====
            
              If LCase(Instrument_Type) = "vhdvs" Then
                   DCVSMeasRangeList = TheHdw.DCVS.pins(Pin_Ary(i)).CurrentRange.list
                   First_measure_max = 0
                   For Each site In TheExec.sites.Selected ' find all sites' max value
                       If First_meas.pins(Pin_Ary(i)).value > First_measure_max Then
                            First_measure_max = First_meas.pins(Pin_Ary(i)).value
                       End If
                   Next site
                   First_measure_max = First_measure_max * 1.5
                   'New_I_Range = thehdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max
                   
                   For j = 0 To UBound(DCVSMeasRangeList)  'compare range list to define new range
                        If DCVSMeasRangeList(j) > First_measure_max Then
                            New_I_Range = DCVSMeasRangeList(j)
                            Exit For
                                     
                        End If
                   Next j
                   
                   If New_I_Range = TheHdw.DCVS.pins(Pin_Ary(i)).CurrentRange.Max Then
                        New_I_Range_Fold = New_I_Range * 0.7
                   Else
                        New_I_Range_Fold = New_I_Range
                   End If
                   
                End If
            '==================================================================================
            
           '+++++++++++++++++++++++++++++++++++++++++++++ apply new current ange +++++++++++++++++++++++++++++++++++++
            With TheHdw.DCVS.pins(Pin_Ary(i))
                   
                    .CurrentRange.value = New_I_Range ' 0.7 ' 0.0024
                    .Meter.CurrentRange = New_I_Range
                    .CurrentLimit.Source.FoldLimit.level.value = New_I_Range_Fold '0.8 * .CurrentLimit.Source.FoldLimit.Level.max
                    .CurrentLimit.Sink.FoldLimit.level.value = 1 * .CurrentLimit.Sink.FoldLimit.level.Max
                    .Gate = True
                    TheHdw.Wait 0.02
            End With
          '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
'''''''  plot current  profile
'''''        While TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.IsRunning
'''''          TheHdw.Wait (0.001)
'''''        Wend
'''''          Current_DSPW = TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.Signals("DCVSCaptureLabel").DSPWave
'''''          Current_DSPW.Plot ("DCVS current")


          Power_data = TheHdw.DCVS.pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
          
         ''''' print current measurement for new curren range
         ' theexec.Datalog.WriteComment ("Current Range : " & New_I_Range & "A")
         ' theexec.Flow.TestLimit resultVal:=Power_data, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitCustom, ScaleType:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=force_v, ForceUnit:=unitVolt, customUnit:="A", ForceResults:=tlForceNone

          Power_volt.AddPin (Pin_Ary(i))
          Power_volt.pins(Pin_Ary(i)).value = TheHdw.pins(Pin_Ary(i)).DCVS.Voltage.Main.value
          Power_data = Power_volt.Math.divide(Power_data)

         
          '---change judge hi/li limit here for different spec 20190807---
          Dim sPin_Key As String
             For Each site In TheExec.sites.Selected
                    sPin_Key = CStr(site) & CStr(Pin_Ary(i))
                    hiLimit = Dict_Res_Hlimit(sPin_Key)
                    
                    'dHiLimit = Dict_Res_Hlimit(sPin)
                    ' TheExec.Flow.TestLimit pldMeasVal, Tname:=strTName, lowVal:=dLoLimit, hiVal:=dHiLimit, unit:=unitCustom, customUnit:=strUnit
                    TheExec.flow.TestLimit resultVal:=Power_data, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=Force_volt, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                    'glb_TestInstance = theexec.DataManager.instancename
                    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
             Next site
            '---change judge hi/li limit here for different spec 20190807---
            Set Power_data = Nothing
            Set Power_volt = Nothing
            
          '********************resistor test end********************'
    
            TheHdw.DCVS.pins(Pin_Ary(i)).Alarm(tlDCVIAlarmMode) = tlAlarmDefault
    
            '***disconnect DCVS instrument***'
            With TheHdw.DCVS.pins(Pin_Ary(i))
                    .CurrentRange.value = 0.02
                    .Voltage.Main = 0#
                    .Voltage.Alt = 0#
                    .Gate = False
                    TheHdw.Wait 0.002
                    .Disconnect tlDCVSConnectDefault
            End With
            TheHdw.Wait 0.0005

   ' Next site
Next i
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Power_Trace_Resistance_Meas"
    If AbortTest Then Exit Function Else Resume Next
End Function

Function UVI80_Power_Trace_Resistance_Meas(TestPin As PinList, RelayOn_Pin As String) As Long
 Dim PowerPins As String
 Dim p As Variant, Pin_Ary() As String, p_cnt As Long
 Dim Res_Value As Double
 Dim DC07_Power_data As New PinListData
 Dim DC07_Power_data_value As Double
 Dim i As Integer
 Dim hiLimit As Double
 Dim LowLimit As Double
 Dim force_i As Double
 Dim force_v As Double
 Dim Tname As String
 Dim sPin As String
 On Error GoTo errHandler
  hiLimit = 10#
LowLimit = 0
force_i = 0.1
force_v = 0.85
Tname = "UVI80 path"


Dim Relay_on_pin_temp1() As String 'add 191121
Dim Relay_on_pin_temp2() As String 'add 191121
Dim Relay_Name As String

Relay_on_pin_temp1 = Split(RelayOn_Pin, ";") 'add 191121
For i = 0 To UBound(Relay_on_pin_temp1)
    Relay_on_pin_temp2 = Split(Relay_on_pin_temp1(i), "+")
    Relay_On_Pin_Dic_short_UVI80(Relay_on_pin_temp2(0)) = Relay_on_pin_temp2(1)
Next i

        TheExec.DataManager.DecomposePinList TestPin, Pin_Ary, p_cnt
        'Pin_dcvi_Ary = Split(p_dc07, ",")
     ' thehdw.DCVI.Pins(AllPowerPins).SetCurrentAndRange 0.02, 0.02
        'FoldLimit = thehdw.DCVI.Pins(AllPowerPins).Current
 For i = 0 To UBound(Pin_Ary)
 sPin = Pin_Ary(i)
 
          If Relay_On_Pin_Dic_short_UVI80.Exists(Pin_Ary(i)) Then
                Relay_Name = Relay_On_Pin_Dic_short_UVI80(Pin_Ary(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOn
                TheHdw.Wait 0.0001
          
          End If
     ' For Each site In theexec.sites.Existing
                 'Below pins are connected to DUT through relays
'''''         If sPin = "ANALOGMUX_OUT" Then
'''''            thehdw.Utility.pins("K08").State = tlUtilBitOn
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "ANALOGMUX_OUT_SRC" Then
'''''            thehdw.Utility.pins("K09,K07").State = tlUtilBitOn
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "PAD_MTR_ANALOG_TEST_P" Then
'''''            thehdw.Utility.pins("K95").State = tlUtilBitOn
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "PAD_MTR_ANALOG_TEST_N" Then
'''''            thehdw.Utility.pins("K96").State = tlUtilBitOn
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "PAD_MTR_VREF_P" Then
'''''            thehdw.Utility.pins("K97").State = tlUtilBitOn
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "PAD_MTR_VREF_N" Then
'''''            thehdw.Utility.pins("K98").State = tlUtilBitOn
'''''            thehdw.Wait 0.0001
'''''         End If
          TheHdw.DCVI.pins(Pin_Ary(i)).Connect tlDCVIConnectDefault
         TheHdw.Wait 0.0005
        
        
         
        
         '***set DCVI instrument***'
         With TheHdw.DCVI.pins(Pin_Ary(i))
            .mode = tlDCVIModeCurrent
            .ComplianceRange(tlDCVICompliancePositive).value = 7
            .ComplianceRange(tlDCVIComplianceNegative).value = 2
             TheHdw.Wait 0.01
            .Current = force_i
            .Voltage = force_v
            .CurrentRange = force_i
            .VoltageRange = force_v
            .NominalBandwidth = 1000
            .Meter.mode = tlDCVIMeterVoltage
            .Meter.Filter.bypass = False
            .Meter.Filter.value = 10000
            .Meter.VoltageRange = force_v
            .Gate = True
            TheHdw.Wait 0.01
         End With
        '***-----------------------***'

        '********************Resistor main test begin********************'
        Dim pldMeasureValue As New PinListData
        Set pldMeasureValue = TheHdw.DCVI.pins(Pin_Ary(i)).Meter.Read(tlStrobe, 100, 100000)
        ' theexec.Flow.TestLimit pldMeasureValue, Tname:=Tname, lowVal:=0, hiVal:=HiLimit, Unit:=8, customUnit:="V"
        pldMeasureValue = pldMeasureValue.Math.divide(TheHdw.pins(Pin_Ary(i)).DCVI.Current)
       '---change judge hi/li limit here for different spec 20190807----
        Dim sPin_Key As String
        'Dim Site As Variant
         ' Dim MeasureValue_site() As New SiteDouble
          
          'MeasureValue_site = pldMeasureValue
         
          For Each site In TheExec.sites.Selected
                sPin_Key = CStr(site) & CStr(Pin_Ary(i))
                hiLimit = Dict_Res_Hlimit(sPin_Key)
                'dHiLimit = Dict_Res_Hlimit(sPin)
                ' TheExec.Flow.TestLimit pldMeasVal, Tname:=strTName, lowVal:=dLoLimit, hiVal:=dHiLimit, unit:=unitCustom, customUnit:=strUnit
                TheExec.flow.TestLimit pldMeasureValue, Tname:=Tname, lowVal:=0, hiVal:=hiLimit, unit:=8, customUnit:="Ohm"
          
          Next site
        '---change judge hi/li limit here for different spec 20190807----
        
    
        
        
        'pldPowerPathResistance.AddPin (Pin_Ary(i))
        'pldPowerPathResistance.Pins(Pin_Ary(i)) = pldMeasureValue.Pins(Pin_Ary(i))
        

        
        Set pldMeasureValue = Nothing
        '********************resistor test end********************'

        TheHdw.DCVI.pins(Pin_Ary(i)).Alarm(tlDCVIAlarmMode) = tlAlarmDefault

        '***disconnect DCVI instrument***'
        With TheHdw.DCVI.pins(Pin_Ary(i))
                .Gate = False
                TheHdw.Wait 0.002
                .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 0.0005
        '***-----------------------***'
        
          If Relay_On_Pin_Dic_short_UVI80.Exists(Pin_Ary(i)) Then
                Relay_Name = Relay_On_Pin_Dic_short_UVI80(Pin_Ary(i))
                TheHdw.Utility.pins(Relay_Name).State = tlUtilBitOn
                TheHdw.Wait 0.0001
          
          End If
'''''       If sPin = "ANALOGMUX_OUT" Then
'''''            thehdw.Utility.pins("K08").State = tlUtilBitOff
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "ANALOGMUX_OUT_SRC" Then
'''''            thehdw.Utility.pins("K09,K07").State = tlUtilBitOff
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "PAD_MTR_ANALOG_TEST_P" Then
'''''            thehdw.Utility.pins("K95").State = tlUtilBitOff
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "PAD_MTR_ANALOG_TEST_N" Then
'''''            thehdw.Utility.pins("K96").State = tlUtilBitOff
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "PAD_MTR_VREF_P" Then
'''''            thehdw.Utility.pins("K97").State = tlUtilBitOff
'''''            thehdw.Wait 0.0001
'''''         ElseIf sPin = "PAD_MTR_VREF_N" Then
'''''            thehdw.Utility.pins("K98").State = tlUtilBitOff
'''''            thehdw.Wait 0.0001
'''''         End If
         TheHdw.Wait 0.001
     ' End If
      

        'Next site
  Next i
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in UNI80_Power_Trace_Resistance_Meas"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function PC_Leakage(ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, leakage_pins As PinList, Hi_Limit As Double, Lo_Limit As Double, Relay_On As PinList)
                           
                

Dim site As Variant
'Dim SeqLeakPins As String
Dim PinArr() As String, PinCount As Long, i As Long
Dim p As Variant
Dim MeasVal As New PinListData
Dim TestNum As Long
Dim Tname As String
Dim AllSitePass As Boolean
Dim BurstResult As New SiteLong

On Error GoTo errHandler

'Call thehdw.Digital.ApplyLevelsTiming(True, True, False, tlPowered, leakage_pins, , leakage_pins)


TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
TheHdw.PPMU.pins(leakage_pins).ForceV (0)
TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
TheHdw.PPMU.pins(leakage_pins).Disconnect

    ''''''use the "theexec.DataManager.DecomposePinList" to serialize the pins to be tested sequentially'''''
    TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
'    DCVS_Trim_NC_Pin PinArr(), PinCount
    
    ''High
    
    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    TheHdw.Digital.pins(leakage_pins).Disconnect
    
    With TheHdw.PPMU(leakage_pins)
  
        .Connect
        .Gate = tlOn
        .ForceV ForceV_IiH, I_Meas_Range
         TheHdw.Wait 0.01
         'DebugPrintFunc_PPMU leakage_pins.Value
         MeasVal = .Read(tlPPMUReadMeasurements)
         
        .ForceV (0)
        .Gate = tlOff
        .Disconnect
    End With
        
    'TheHdw.digital.Pins(leakage_pins).Connect 'connect the tested pin back to the PE
    
    'offline mode simulation
    If TheExec.TesterMode = testModeOffline Then

        For Each site In TheExec.sites
            For Each p In PinArr()
                If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.pins(p).value = 5 * uA + Rnd() * 0.1 * uA
            Next p
        Next site
    End If
  
  
    TheExec.flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:="Leakage", ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    'TheExec.Flow.TestLimit resultval:=PPMUMeasure.Pins(DUTPin), lowval:=LowLimit, hival:=HiLimit, ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", tname:=tname, forceVal:=force_i, forceunit:=unitAmp, forceResults:=tlForceNone

TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
   


Exit Function
errHandler:
    TheExec.AddOutput "Error in the Seq Leakage Test"
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Loopback_IO_Leakage(ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, leakage_pins As PinList, Hi_Limit As Double, Lo_Limit As Double, Relay_On As PinList)
                           
                

Dim site As Variant
'Dim SeqLeakPins As String
Dim PinArr() As String, PinCount As Long, i As Long
Dim p As Variant
Dim MeasVal As New PinListData
Dim TestNum As Long
Dim Tname As String
Dim AllSitePass As Boolean
Dim BurstResult As New SiteLong
Dim Tname_str As String

On Error GoTo errHandler

'Call thehdw.Digital.ApplyLevelsTiming(True, True, False, tlPowered, leakage_pins, , leakage_pins)
Tname_str = "Loopback Pin Leakage Meas"

TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
TheHdw.PPMU.pins(leakage_pins).ForceV (0)
TheHdw.PPMU.pins(leakage_pins).Gate = tlOff
TheHdw.PPMU.pins(leakage_pins).Disconnect

    ''''''use the "theexec.DataManager.DecomposePinList" to serialize the pins to be tested sequentially'''''
    TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
'    DCVS_Trim_NC_Pin PinArr(), PinCount
    
    ''High
    
    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    TheHdw.Digital.pins(leakage_pins).Disconnect
For Each p In PinArr
    With TheHdw.PPMU(p)
  
        .Connect
        .Gate = tlOn
        .ForceV ForceV_IiH, I_Meas_Range
         TheHdw.Wait 0.01
         'DebugPrintFunc_PPMU leakage_pins.Value
         MeasVal = .Read(tlPPMUReadMeasurements)
         
        .ForceV (0)
        .Gate = tlOff
        .Disconnect
         TheExec.flow.TestLimit resultVal:=MeasVal, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Tname_str, ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
         'glb_TestInstance = theexec.DataManager.instancename
         'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    End With
Next p
    'TheHdw.digital.Pins(leakage_pins).Connect 'connect the tested pin back to the PE
    
    'offline mode simulation
'''''    If TheExec.TesterMode = testModeOffline Then
'''''
'''''        For Each Site In TheExec.sites
'''''            For Each p In PinArr()
'''''                If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.pins(p).Value = 5 * uA + Rnd() * 0.1 * uA
'''''            Next p
'''''        Next Site
'''''    End If
  
  
   ' TheExec.Flow.TestLimit resultVal:=MeasVal, lowVal:=lo_limit, hiVal:=hi_limit, ScaleType:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:="Leakage", forceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
    'TheExec.Flow.TestLimit resultval:=PPMUMeasure.Pins(DUTPin), lowval:=LowLimit, hival:=HiLimit, ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", tname:=tname, forceVal:=force_i, forceunit:=unitAmp, forceResults:=tlForceNone

TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
   


Exit Function
errHandler:
    TheExec.AddOutput "Error in the Seq Leakage Test"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function UVS256_HexVs_Leakage(ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, Power_pins As PinList, Hi_Limit As Double, Lo_Limit As Double)
                          
                
    Dim PinArr() As String, PinCount As Long
    
    Dim sAllDCVSPins As String
    Dim strAllDCVSPins()  As String
    Dim strAllDCVSMerged2Pins()  As String
    Dim strAllDCVSMerged4Pins()  As String
    Dim strAllDCVSMerged8Pins()  As String
    Dim lngPnum As Long
    'Dim arrsPinList() As String
    Dim Hexvs_Pins As String
    Dim i As Integer
    Dim sPin As Variant
    Dim lIndex As Long
    Dim TestName As String
    Dim Min_Current_Range As Double
    Dim Max_Current_Range As Double
    'dim
    Dim sWaveSignalName As String
    Dim dspWave_ICapture As New DSPWave
      sWaveSignalName = "DIBC_DCVS_CapMeasureSig"
'''''    Call Theexec.DataManager.GetPinNames(strAllDCVSPins, chDCVS, lngPnum)
'''''    Call Theexec.DataManager.GetPinNames(strAllDCVSMerged2Pins, chDCVSMerged2, lngPnum)
'''''    Call Theexec.DataManager.GetPinNames(strAllDCVSMerged4Pins, chDCVSMerged4, lngPnum)
'''''    Call Theexec.DataManager.GetPinNames(strAllDCVSMerged8Pins, chDCVSMerged8, lngPnum)
'''''
'''''    sAllDCVSPins = Join(strAllDCVSPins, ",") & "," & Join(strAllDCVSMerged2Pins, ",") & "," & Join(strAllDCVSMerged4Pins, ",") & "," & Join(strAllDCVSMerged8Pins, ",")
'''''
'''''    arrsPinList = Split(sAllDCVSPins, ",")
        
    lIndex = 0
    TheExec.DataManager.DecomposePinList Power_pins, PinArr(), PinCount
    If UCase(Power_pins) Like UCase("*HexVs*") Then
             TestName = "HexVs Leakage"
             For Each sPin In PinArr
                    TheHdw.DCVS.pins(sPin).Connect tlDCVSConnectDefault
                    
                     TheHdw.Wait 0.01
                     With TheHdw.DCVS.pins(sPin)
                        .mode = tlDCVSModeVoltage
                        .Voltage.Main.value = ForceV_IiH
                        .Voltage.Alt.value = ForceV_IiH
                        .Voltage.Output = tlDCVSVoltageMain
                        .CurrentRange.value = I_Meas_Range
                    Min_Current_Range = TheHdw.DCVS.pins(sPin).Meter.CurrentRange.Min
                        .Meter.mode = tlDCVSMeterCurrent
                        .Meter.CurrentRange = Min_Current_Range
                        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
                        .Gate = True
                        TheHdw.Wait 1#
                      End With
                     '***-----------------------***'
                      TestName = sPin & " pin Leakage"
                      Dim pldMeasureValue As New PinListData
                     '********************main test begin********************'
                      pldMeasureValue = TheHdw.DCVS.pins(sPin).Meter.Read(tlStrobe, 1000, 100000, tlDCVIMeterReadingFormatAverage)
                     '********************main test end********************'
            
                     '********************Power Supply leakage test end********************'
                    TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TestName, ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                    'glb_TestInstance = theexec.DataManager.instancename
                    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
''                    judgeCheckResult pldMeasureValue, 13600 + lIndex, "DCVS_Short"
''                    If Theexec.sites.ActiveCount = 0 Then Exit Sub
''
''                    lIndex = lIndex + 1
        
            Next sPin
    ElseIf UCase(Power_pins) Like UCase("*VS256*") Then
             
            TestName = "UVS256 Leakage"
           ' Dim sAllDCVSPins As String
        '   Dim lIndex As Long
        '   tl_PinListDataSort (False) '''' set the measure pin sequence
           
           Call TheExec.DataManager.GetPinNames(strAllDCVSPins, chDCVS, lngPnum)
           Call TheExec.DataManager.GetPinNames(strAllDCVSMerged2Pins, chDCVSMerged2, lngPnum)
           Call TheExec.DataManager.GetPinNames(strAllDCVSMerged4Pins, chDCVSMerged4, lngPnum)
           Call TheExec.DataManager.GetPinNames(strAllDCVSMerged8Pins, chDCVSMerged8, lngPnum)
           
           sAllDCVSPins = Join(strAllDCVSPins, ",") & "," & Join(strAllDCVSMerged2Pins, ",") & "," & Join(strAllDCVSMerged4Pins, ",") & "," & Join(strAllDCVSMerged8Pins, ",")
           strAllDCVSPins = Split(sAllDCVSPins, ",")

                     
            TheHdw.DCVS.pins(sPin).Connect tlDCVSConnectDefault
            TheHdw.Wait 0.1
            For Each sPin In strAllDCVSPins
            
            
            
                 Max_Current_Range = TheHdw.DCVS.pins(sPin).CurrentRange.Max
                If UCase(sPin) = "SPI_ROM_PWR" Or UCase(sPin) = "BUFFER_PWR_2_DOMAIN" Or UCase(sPin) = "VDD_AMPH" Or UCase(sPin) = "BUFFER_PWR" Then 'conenct to IC
                ElseIf UCase(sPin) = "VDD_GPU" Or UCase(sPin) = "VDD_SOC" Or UCase(sPin) = "VDD_PCPU" Or UCase(sPin) = "VDD1" Or UCase(sPin) = "VDD2" Or UCase(sPin) = "VDD_ECPU" Or UCase(sPin) = "VDD_CPU_SRAM" Then
                    GoTo skip 'will execute later
                Else 'If UCase(sPin) = "VDD12_DN_PCIE" Then

'''''                 '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'''''                       With thehdw.DCVS.Pins(sPin).Capture
'''''                            .Signals.Add (sWaveSignalName)
'''''                            .Signals(sWaveSignalName).mode = tlDCVSMeterCurrent
'''''                            .Signals(sWaveSignalName).SampleRate = 200000#
'''''                            .Signals(sWaveSignalName).SampleSize = 16384#
'''''                            .Signals(sWaveSignalName).LoadSettings
'''''                        End With
'''''
'''''                        Call thehdw.Wait(0.001)
'''''                         ' Trigger the capture
'''''                        Call thehdw.DCVS.Pins(sPin).Capture.Signals(sWaveSignalName).Trigger
                 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '  For Each Site In TheExec.Sites.Selected
                    TheHdw.DCVS.pins(sPin).Connect tlDCVSConnectDefault
  
                     TheHdw.Wait 0.1
                     With TheHdw.DCVS.pins(sPin)
                        .mode = tlDCVSModeVoltage
                    
                        .Voltage.Main.value = ForceV_IiH
                        .Voltage.Alt.value = ForceV_IiH
                        .Voltage.Output = tlDCVSVoltageMain
                        .CurrentRange.value = 0.2 '.02
                        '.CurrentRange.Value = Max_Current_Range
                        .Meter.mode = tlDCVSMeterCurrent
                        .Meter.CurrentRange = 0.2
                        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
                    
                        .Gate = True


                       TheHdw.Wait 0.4
'
'
                        .CurrentRange.value = I_Meas_Range '0.000004
                        .Meter.CurrentRange = I_Meas_Range '0.000004

                        TheHdw.Wait 0.4
                      End With
                     '***-----------------------***'
                      TestName = sPin & " pin Leakage"
                    'Dim pldMeasureValue As New PinListData
                    Dim m_value As Double
                     '********************main test begin********************'
'''''                       dspWave_ICapture = thehdw.DCVS.Pins(sPin).Capture.Signals(sWaveSignalName).DSPWave
                     pldMeasureValue = TheHdw.DCVS.pins(sPin).Meter.Read(tlStrobe, 1000, 100000, tlDCVIMeterReadingFormatAverage)
                    ' M_value = thehdw.DCVS.Pins(sPin).Meter.Read(tlStrobe, 1000, 100000, tlDCVIMeterReadingFormatAverage)
                     '********************main test end********************'
            
                     '********************Power Supply leakage test end********************'
                       TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TestName, ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                       'glb_TestInstance = theexec.DataManager.instancename
                       'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                        'TheExec.Flow.TestLimit resultVal:=M_value, lowVal:=lo_limit, hiVal:=hi_limit, ScaleType:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=TestName, forceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone

''''
''''                   '***disconnect DCVS instrument***'
''''    With thehdw.DCVS.Pins(sPin)
''''       .CurrentRange.Value = 0.02
''''       .Meter.CurrentRange = 0.02
''''       .Voltage.Main = 0#
''''       .Voltage.Alt = 0#
''''       .Gate = False
''''       thehdw.Wait 0.002
''''       .Disconnect tlDCVSConnectDefault
''''    End With
''''    thehdw.Wait 0.0005
''''    '***-----------------------***'

'                    judgeCheckResult pldMeasureValue, 13600 + lIndex, "DCVS_Short"
'                    If Theexec.sites.ActiveCount = 0 Then Exit Sub
'
'                    lIndex = lIndex + 1
                 '  Next Site
                End If
skip:
            Next sPin
    End If
    Set pldMeasureValue = Nothing
    
    '***disconnect DCVS instrument***'
    With TheHdw.DCVS.pins(Power_pins)
       .CurrentRange.value = 0.02
       .Meter.CurrentRange = 0.02
       .Voltage.Main = 0#
       .Voltage.Alt = 0#
       .Gate = False
       TheHdw.Wait 0.002
       .Disconnect tlDCVSConnectDefault
    End With
    TheHdw.Wait 0.0005
    '***-----------------------***'
      
   

    Set pldMeasureValue = Nothing


Exit Function
errHandler:
    TheExec.AddOutput "Error in the Seq Leakage Test"
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function UVI80_Leakage(ForceV_IiH As Double, ForceV_IiL As Double, I_Meas_Range As Double, Power_pins As PinList, Hi_Limit As Double, Lo_Limit As Double)
                          
                
    Dim PinArr() As String, PinCount As Long
    
    Dim strAllDCVIPins() As String
    Dim sAllDCVIPins As String
    Dim lngPnum As Long
    Dim sPin As Variant
    Dim vsite As Variant
    Dim pldMeasureValue As New PinListData
    Dim TestName As String
'    Dim pldMeasureValue_temp As New PinListData
    
    Dim lIndex As Long
    
'''    Call Theexec.DataManager.GetPinNames(strAllDCVIPins, chDCVI, lngPnum)
'''    sAllDCVIPins = Join(strAllDCVIPins, ",")
    
    TheExec.DataManager.DecomposePinList Power_pins, PinArr(), PinCount
    For Each sPin In PinArr
        If UCase(sPin) = "ANALOGMUX_OUT_GND" Or UCase(sPin) = "HSC_VBG_VSS" Then
        
        Else
            '***connect DCVI instrument***'
            TheHdw.DCVI.pins(sPin).Connect tlDCVIConnectDefault
            TheHdw.DCVI.pins(sPin).LocalKelvin = True
            TheHdw.Wait 0.0005
            '***-----------------------***'

               '***set DCVI instrument***'
            With TheHdw.DCVI.pins(sPin)
               .mode = tlDCVIModeVoltage
               .ComplianceRange(tlDCVICompliancePositive).value = 7
               .ComplianceRange(tlDCVIComplianceNegative).value = 2
               TheHdw.Wait 0.01
               .Voltage = ForceV_IiH
               .Current = I_Meas_Range
               .VoltageRange.value = 2
               .CurrentRange.value = I_Meas_Range
               .NominalBandwidth = 1003
               .Meter.mode = tlDCVIMeterCurrent
               .Meter.CurrentRange = I_Meas_Range
               .Meter.Filter.bypass = False
               .Meter.Filter.value = 10000
               .Alarm(tlDCVIAlarmMode) = tlAlarmOff
               .Gate = True
               TheHdw.Wait 0.01
               .SetCurrentAndRange 20 * uA, 20 * uA
               TheHdw.Wait 0.001
            End With
               '***-----------------------***'
            TestName = sPin & " pin Leakage"
            '********************main test begin********************'
            pldMeasureValue = TheHdw.DCVI.pins(sPin).Meter.Read(tlStrobe, 10, 100000, tlDCVIMeterReadingFormatAverage)
            TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=Lo_Limit, hiVal:=Hi_Limit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TestName, ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            'glb_TestInstance = theexec.DataManager.instancename
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
''                       If Theexec.sites.ActiveCount = 0 Then Exit Sub
''
''            lIndex = lIndex + 1
             
            '***disconnect DCVI instrument***'
            With TheHdw.DCVI.pins(sPin)
            .SetCurrentAndRange 0.01, 0.01
               .Gate = False
               TheHdw.Wait 0.002
               .Reset tlResetSettings + tlResetConnections
            End With
            TheHdw.Wait 0.0005
        
          End If
             
        Next sPin
               '********************main test end********************'
                
               '********************path test end********************'
       ' Set pldMeasureValue = Nothing
            
       TheHdw.DCVI.pins(Power_pins).Alarm(tlDCVIAlarmMode) = tlAlarmDefault

Exit Function
errHandler:
    TheExec.AddOutput "Error in the Seq Leakage Test"
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function Relay_Circuit_check(R_LowLimit As Double, R_HiLimit As Double, Force_volt As Double, Current_Range As Double, Off_V_LowLimit As Double, Off_V_HiLimit As Double, Optional Relay_On As PinList, Optional WaitTime As Double = 0.003, Optional Route_Start As PinList, Optional Route_End As PinList) As Long

'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim Wait_Time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    'Dim R_Volt As PinListData
    Dim R_Value As New PinListData
    Dim F_volt As New PinListData
    Dim Res_Value As Double
    Dim FVolt As Double
    Dim CurrValue As Double
    'Dim FCurr As Double
'''    Dim HiLimit As Double
'''    Dim LowLimit As Double
    On Error GoTo errHandler
''''    HiLimit = 300#
''''    LowLimit = 200#
    'FVolt = 0.02
    'FCurr = 0.0004
    relayOnStr = ""
    relayOffStr = ""
  
    TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On

    
'''    Trim_NC_Pin Pins_On, Pin_Cnt_On
'''    Trim_NC_Pin Pins_Off, Pin_Cnt_Off


   
   For Each site In TheExec.sites.Selected
     If Relay_On <> "" Then
        'theexec.Datalog.WriteComment "============================ Relay ON ====================================="
        Tname = "Site" & site & "-" & CStr(Relay_On) & " On"
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
        For Each p In Pins_On
                'Tname = "rly_on_" & p
                BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)
    
            'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
           ' TheExec.Datalog.WriteComment "Relay_on : " & BitState.Pins(p)
        Next p
    Else
        Tname = TheExec.DataManager.instancename
     End If
          If Route_End <> "" Then
                If LCase(TheExec.DataManager.pinType(Route_End)) = "power" Then
                
                        If LCase(TheExec.DataManager.ChannelType(Route_End)) Like "*dcvs*" Then
                         TheHdw.DCVS.pins(Route_End).Connect tlDCVSConnectDefault
                         TheHdw.Wait 0.0005
                        '***-----------------------***'
                                    With TheHdw.DCVS.pins(Route_End)
                                       .mode = tlDCVSModeVoltage
                                       .Voltage.Main.value = 0
                                       .Voltage.Alt.value = 0
                                       .Voltage.Output = tlDCVSVoltageMain
                                       .CurrentRange.value = 0.002
                                       .Meter.mode = tlDCVSMeterCurrent
                                       .Meter.CurrentRange = 0.002
                                       .Alarm(tlDCVSAlarmAll) = tlAlarmOff
                                       .Gate = True
                                       TheHdw.Wait 0.25
                                    End With
                        ElseIf LCase(TheExec.DataManager.ChannelType(Route_End)) = "dcvi" Then
                              With TheHdw.DCVI.pins(Route_End)
                                     .mode = tlDCVIModeCurrent
                                     .ComplianceRange(tlDCVICompliancePositive).value = 7
                                     .ComplianceRange(tlDCVIComplianceNegative).value = 2
                                     TheHdw.Wait 0.01
                                     .Current = 0
                                     .Voltage = 0
                                     .CurrentRange = 0.02
                                     .VoltageRange = 2
                                     .NominalBandwidth = 1000
                                     .Meter.mode = tlDCVIMeterVoltage
                                     .Meter.Filter.bypass = False
                                     .Meter.Filter.value = 10000
                                     .Meter.VoltageRange = 2
                                     .Gate = True
                                     TheHdw.Wait 0.01
                             End With
                        End If
                ElseIf LCase(TheExec.DataManager.pinType(Route_End)) = "i/o" Then
                    TheHdw.Digital.pins(Route_End).Disconnect
                    With TheHdw.PPMU.pins(Route_End)
                        .Connect
                        .ForceV 0#
                        .Gate = tlOn
                     
                    End With
                End If
                TheHdw.Wait 0.1
          Else
                
               
          End If
          
          
          If LCase(TheExec.DataManager.ChannelType(Route_Start)) = "i/o" Then
                    TheHdw.Digital.pins(Route_Start).Disconnect
                    With TheHdw.PPMU.pins(Route_Start)
                        .Connect
                        '.ForceI 0.002, 0.002
                        .ForceV Force_volt, Current_Range
                        .Gate = tlOn
                        
                    End With
                
                     TheHdw.Wait 0.1
                   '----------------------------- Calculate Measure result
                   
                    R_Value = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
                   ' R_Value = R_Value.Math.Divide(thehdw.PPMU.Pins(Route_Start).current.Value)
                    'F_volt.AddPin(Route_Start).Value = thehdw.Pins(Route_Start).PPMU.Voltage.Value
                    F_volt.AddPin (Route_Start)
                    F_volt.pins(Route_Start).value = TheHdw.pins(Route_Start).PPMU.Voltage.value
                    If R_Value = 0 Then
                        R_Value = 99999999
                    Else
                        R_Value = F_volt.Math.divide(R_Value)
                    End If
                    TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=Force_volt, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                    'glb_TestInstance = theexec.DataManager.instancename
                    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
           ElseIf LCase(TheExec.DataManager.ChannelType(Route_Start)) = "dcvi" Then
                        '***connect DCVI instrument***'
                    TheHdw.DCVI.pins(Route_Start).Connect tlDCVIConnectDefault
                    TheHdw.DCVI.pins(Route_Start).Disconnect tlDCVIConnectHighSense
                  '  thehdw.DCVI.pins(Route_Start).Disconnect tlDCVIConnectHighForce
                    TheHdw.DCVI.pins(Route_Start).LocalKelvin = True
                    TheHdw.Wait 0.0005
                    '***-----------------------***'
                                        '***set DCVI instrument***'
                    
                    '***-----------------------***'
            
                    TheHdw.Wait 0.003
            
                    '********************resistor main test begin********************'
'''                    Dim dblMeasureCurrent As Double
'''                    dblMeasureCurrent = 0.004
            
                    '***set DCVI instrument***'
                    With TheHdw.DCVI.pins(Route_Start)
                            .mode = tlDCVIModeCurrent
                            .ComplianceRange(tlDCVICompliancePositive).value = 7
                            .ComplianceRange(tlDCVIComplianceNegative).value = 2
                            TheHdw.Wait 0.01
                            .Current = Current_Range
                            .Voltage = Force_volt '1000 * 1.5 * Current_Range
                            .CurrentRange = Current_Range
                            .VoltageRange = 1000 * 1.3 * Current_Range
                            .NominalBandwidth = 1000
                            .Meter.mode = tlDCVIMeterVoltage
                            .Meter.Filter.bypass = False
                            .Meter.Filter.value = 10000
                            .Meter.VoltageRange = 1000 * 1.2 * Current_Range
                            .Gate = True
                            TheHdw.Wait 0.01
                    End With
                    '***-----------------------***'

                '********************Resistor main test begin********************'
                'Dim pldMeasureValue As New PinListData
                R_Value = TheHdw.DCVI.pins(Route_Start).Meter.Read(tlStrobe, 100, 100000)
                CurrValue = TheHdw.pins(Route_Start).DCVI.Current
                If CurrValue = 0 Then
                    R_Value = 99999999
                Else
                    R_Value = R_Value.Math.divide(CurrValue)
                End If
                TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=Force_volt, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                'glb_TestInstance = theexec.DataManager.instancename
                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
           End If
           
           
            Set R_Value = Nothing
            Set F_volt = Nothing
           '----------------------------------------------------------------------------------------------------------------------------------------------------------
      
            If Relay_On <> "" Then

                
                
                Tname = "Site" & site & "-" & CStr(Relay_On) & " Off"
                TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
                '''TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
                
                For Each p In Pins_On
                        'Tname = "rly_on_" & p
                        BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)
                
                        'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
                       ' TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
                Next p
                If LCase(TheExec.DataManager.ChannelType(Route_Start)) = "i/o" Then
                             TheHdw.PPMU.pins(Route_Start).ClampVHi = 1#
                             
                             TheHdw.PPMU.pins(Route_Start).ForceV 0#
                             
                             TheHdw.PPMU.pins(Route_Start).ForceV Force_volt, Current_Range
                        
                             TheHdw.Wait 0.002
                             R_Value = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
                             
                             Dim counter As Integer: counter = 0
                             
                             While R_Value.Math.compare(EqualTo, 0) And counter <= 10
                               R_Value = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
                               counter = counter + 1
                             Wend
                             
                             F_volt.AddPin (Route_Start)
                             F_volt.pins(Route_Start).value = TheHdw.pins(Route_Start).PPMU.Voltage.value
                             
                            If R_Value = 0 Then
                                R_Value = 99999999
                            Else
                                R_Value = F_volt.Math.divide(Abs(R_Value))
                            End If
                
                             
                             TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=Force_volt, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
                             'glb_TestInstance = theexec.DataManager.instancename
                             'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                             Set R_Value = Nothing
                             Set F_volt = Nothing
                             With TheHdw.PPMU.pins(Route_Start)
                                 .ForceV (0)
                                 .Disconnect
                                 .Gate = tlOff
                         
                             End With
                  ElseIf LCase(TheExec.DataManager.ChannelType(Route_Start)) = "dcvi" Then
                                             '***set DCVI instrument***'
                            With TheHdw.DCVI.pins(Route_Start)
                                    .Current = 0.0015
                                    .Voltage = 1
                                    .CurrentRange = 0.002
                                    .VoltageRange = 1
                                    .Meter.VoltageRange = 2
                                    .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                                    .Gate = False
                                    TheHdw.Wait 0.005
                            End With
                            '***-----------------------***'
                    
                            TheHdw.Wait 0.003
                    
                            TheHdw.DCVI.pins(Route_Start).Gate = True
                            TheHdw.Wait 0.003
                            '********************subtest begin 1********************'
                            R_Value = TheHdw.DCVI.pins(Route_Start).Meter.Read(tlStrobe, 100, 100000)
                            TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, hiVal:=Off_V_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=Force_volt, ForceUnit:=unitAmp, customUnit:="V", ForceResults:=tlForceNone
                            'glb_TestInstance = theexec.DataManager.instancename
                            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                  End If
                TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
                
            End If
            If Route_End <> "" Then
                If LCase(TheExec.DataManager.pinType(Route_End)) = "power" Then
                     With TheHdw.DCVS.pins(Route_End)
                         .mode = tlDCVSModeVoltage
                         .Voltage.Main = 0#
                         .Disconnect
                         .Gate = False
                     End With
                 ElseIf LCase(TheExec.DataManager.pinType(Route_End)) = "i/o" Then
                     TheHdw.Digital.pins(Route_End).Disconnect
                     With TheHdw.PPMU.pins(Route_End)
                         .ForceV 0#
                         .Disconnect
                         .Gate = tlOff
                         
                     End With
                End If
             Else
                 
             End If
            
              
   Next site
    TheHdw.Digital.pins("All_Digital").Disconnect
    TheHdw.PPMU.pins("All_Digital").Disconnect
Exit Function


errHandler:
    ErrorDescription ("Relay_Resistance")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Relay_Circuit_check_short(R_LowLimit As Double, R_HiLimit As Double, Force_volt As Double, Current_Range As Double, Off_V_LowLimit As Double, Off_V_HiLimit As Double, Optional Relay_On As PinList, Optional WaitTime As Double = 0.003, Optional Route_Start As PinList, Optional Route_End As PinList) As Long

'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim Wait_Time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    'Dim R_Volt As PinListData
    Dim R_Value As New PinListData
    Dim F_volt As New PinListData
    Dim Res_Value As Double
    Dim FVolt As Double
    Dim i As Integer
    'Dim FCurr As Double
'''    Dim HiLimit As Double
'''    Dim LowLimit As Double
    On Error GoTo errHandler
''''    HiLimit = 300#
''''    LowLimit = 200#
    'FVolt = 0.02
    'FCurr = 0.0004
    relayOnStr = ""
    relayOffStr = ""
  
    TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On
    TheExec.DataManager.DecomposePinList Route_Start, Pin_Ary, Pin_Cnt
    
'''    Trim_NC_Pin Pins_On, Pin_Cnt_On
'''    Trim_NC_Pin Pins_Off, Pin_Cnt_Off



   For Each site In TheExec.sites.Selected
        For i = 0 To UBound(Pin_Ary)
          If Relay_On <> "" Then
             'theexec.Datalog.WriteComment "============================ Relay ON ====================================="
             Tname = "Site" & site & "-" & CStr(Relay_On) & " Off"
             TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
             For Each p In Pins_On
                     'Tname = "rly_on_" & p
                     BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)
         
                 'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
                ' TheExec.Datalog.WriteComment "Relay_on : " & BitState.Pins(p)
             Next p
          End If
               If Route_End <> "" Then
                     If LCase(TheExec.DataManager.pinType(Route_End)) = "power" Then
                     
                         With TheHdw.DCVS.pins(Route_End)
                             .Connect
                             .mode = tlDCVSModeVoltage
                             .Voltage.Main = 0#
                             .Gate = True
                          
                         End With
                     ElseIf LCase(TheExec.DataManager.pinType(Route_End)) = "i/o" Then
                         TheHdw.Digital.pins(Route_End).Disconnect
                         With TheHdw.PPMU.pins(Route_End)
                             .Connect
                             .ForceV 0#
                             .Gate = tlOn
                          
                         End With
                     End If
              Else
                     
                    
             End If
                 TheHdw.Digital.pins(Pin_Ary(i)).Disconnect
                 With TheHdw.PPMU.pins(Pin_Ary(i))
                     .Connect
                     '.ForceI 0.002, 0.002
                     .ForceV Force_volt, Current_Range
                     .Gate = tlOn
                     
                 End With
             
                  TheHdw.Wait 0.1
                '----------------------------- Calculate Measure result
                
                 R_Value = TheHdw.PPMU.pins(Pin_Ary(i)).Read(tlPPMUReadMeasurements, 10)
                ' R_Value = R_Value.Math.Divide(thehdw.PPMU.Pins(Route_Start).current.Value)
                 'F_volt.AddPin(Route_Start).Value = thehdw.Pins(Route_Start).PPMU.Voltage.Value
                 F_volt.AddPin (Pin_Ary(i))
                 F_volt.pins(Pin_Ary(i)).value = TheHdw.pins(Pin_Ary(i)).PPMU.Voltage.value
                 R_Value = F_volt.Math.divide(R_Value)
                 TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=Force_volt, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                 'glb_TestInstance = theexec.DataManager.instancename
                 'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                 Set R_Value = Nothing
                 Set F_volt = Nothing
                '----------------------------------------------------------------------------------------------------------------------------------------------------------
           
                 If Relay_On <> "" Then
                     TheHdw.PPMU.pins(Pin_Ary(i)).ClampVHi = 1#
                     
                     TheHdw.PPMU.pins(Pin_Ary(i)).ForceV 0#
                     
                     
                     Tname = "Site" & site & "-" & CStr(Relay_On) & " On"
                     TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
                     '''TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
                     
                     For Each p In Pins_On
                             'Tname = "rly_on_" & p
                             BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)
                     
                             'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
                            ' TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
                     Next p
                     
                     TheHdw.PPMU.pins(Pin_Ary(i)).ForceV Force_volt, Current_Range
                
                     TheHdw.Wait 0.1
                     R_Value = TheHdw.PPMU.pins(Pin_Ary(i)).Read(tlPPMUReadMeasurements, 10)
                     F_volt.AddPin (Pin_Ary(i))
                     F_volt.pins(Pin_Ary(i)).value = TheHdw.pins(Pin_Ary(i)).PPMU.Voltage.value
                     R_Value = F_volt.Math.divide(Abs(R_Value))
                     
                     TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=Force_volt, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                     'glb_TestInstance = theexec.DataManager.instancename
                     'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                     Set R_Value = Nothing
                     Set F_volt = Nothing
                     With TheHdw.PPMU.pins(Pin_Ary(i))
                         .ForceV (0)
                         .Disconnect
                         .Gate = tlOff
                 
                     End With
                     TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
                     
                 End If
                 If Route_End <> "" Then
                     If LCase(TheExec.DataManager.pinType(Route_End)) = "power" Then
                          With TheHdw.DCVS.pins(Route_End)
                              .mode = tlDCVSModeVoltage
                              .Voltage.Main = 0#
                              .Disconnect
                              .Gate = False
                          End With
                      ElseIf LCase(TheExec.DataManager.pinType(Route_End)) = "i/o" Then
                          TheHdw.Digital.pins(Route_End).Disconnect
                          With TheHdw.PPMU.pins(Route_End)
                              .ForceV 0#
                              .Disconnect
                              .Gate = tlOff
                              
                          End With
                     End If
                  Else
                      
                  End If
                 
        Next i
   Next site

Exit Function


errHandler:
    ErrorDescription ("Relay_Resistance")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function DPDT_Relay_check(R_LowLimit As Double, R_HiLimit As Double, FCurr As Double, Off_V_LowLimit As Double, Off_V_HiLimit As Double, Optional Relay_Checked As PinList, Optional WaitTime As Double = 0.003, Optional Default_pin As PinList, Optional Switch_pin As PinList)

'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim Wait_Time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    'Dim R_Volt As PinListData
    Dim R_Value As New PinListData
    Dim Res_Value As Double
    Dim FVolt As Double
    Dim i As Integer
    Dim PinArr_D() As String
    Dim PinCount_D As Long
    Dim PinArr_S() As String
    Dim PinCount_S As Long
    'Dim FCurr As Double
'''    Dim HiLimit As Double
'''    Dim LowLimit As Double
    On Error GoTo errHandler
''''    HiLimit = 300#
''''    LowLimit = 200#
    'FVolt = 0.02
    'FCurr = 0.0004
    
    TheHdw.DCVS.pins("All_power").Connect tlDCVSConnectDefault
                     TheHdw.Wait 0.01
                     With TheHdw.DCVS.pins("All_power")
                        .mode = tlDCVSModeVoltage
                        .Voltage.Main.value = 0#
                        .Voltage.Alt.value = 0#
                        .Voltage.Output = tlDCVSVoltageMain
                        .CurrentRange.value = 0.02
                        .Meter.mode = tlDCVSMeterCurrent
                        .Meter.CurrentRange = 0.02
                        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
                        .Gate = True
                        TheHdw.Wait 1#
                      End With
    
    
    
    
    relayOnStr = ""
    relayOffStr = ""
  
    TheExec.DataManager.DecomposePinList Relay_Checked, Pins_On(), Pin_Cnt_On

    
'''    Trim_NC_Pin Pins_On, Pin_Cnt_On
'''    Trim_NC_Pin Pins_Off, Pin_Cnt_Off
    
    'theexec.Datalog.WriteComment "============================ Relay ON ====================================="
    Tname = CStr(Relay_Checked) & " Off"
    TheHdw.Utility.pins(Relay_Checked).State = tlUtilBitOff
    For Each p In Pins_On
            'Tname = "rly_on_" & p
            BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)

        'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
        TheExec.Datalog.WriteComment "Relay_off : " & BitState.pins(p)
    Next p
   TheExec.DataManager.DecomposePinList Default_pin, PinArr_D, PinCount_D
   For i = 0 To UBound(PinArr_D)
           For Each site In TheExec.sites.Existing
                    TheHdw.Digital.pins(PinArr_D(i)).Disconnect
                    With TheHdw.PPMU.pins(PinArr_D(i))
                        .Connect
                        .ForceV FCurr, 0.04
                        .Gate = tlOn
                       
                    End With
                
                     TheHdw.Wait 0.001
                    'R_curr = thehdw.DCVS.Pins(rout_start).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                    R_Value = TheHdw.PPMU.pins(PinArr_D(i)).Read(tlPPMUReadMeasurements, 10)
                    R_Value = TheHdw.PPMU.pins(PinArr_D(i)).Voltage.value / TheHdw.PPMU.pins(PinArr_D(i)).Read(tlPPMUReadMeasurements, 10)
                    'R_Value = R_Value.Math.Divide(thehdw.PPMU.Pins(PinArr_D(i)).Voltage.Value)
                    'Res_Value = R_curr / FCurr
                            ''Res_Value = FVolt / R_curr
                    TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
                    'glb_TestInstance = theexec.DataManager.instancename
                    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                    With TheHdw.PPMU.pins(PinArr_D(i))
                        .ForceV (0)
                        .Disconnect
                        .Gate = tlOff
                
                    End With
           Next site
     Next i
     If Switch_pin <> "" Then
         TheExec.DataManager.DecomposePinList Switch_pin, PinArr_S, PinCount_S
         For i = 0 To UBound(PinArr_S)
                For Each site In TheExec.sites.Existing
                         TheHdw.Digital.pins(PinArr_S(i)).Disconnect
                         TheHdw.PPMU.pins(PinArr_S(i)).ClampVHi = 1#
                           With TheHdw.PPMU.pins(PinArr_S(i))
                             .Connect
                             .ForceI 0.002, 0.002
                             .Gate = tlOn
                             
                         End With
                         'thehdw.PPMU.Pins(PinArr_S(i)).ForceV 0#
                         'thehdw.PPMU.Pins(PinArr_S(i)).ForceI 0.0002, 0.002
                        
                         R_Value = TheHdw.PPMU.pins(PinArr_S(i)).Read(tlPPMUReadMeasurements, 10)
                         TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, hiVal:=Off_V_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="V", ForceResults:=tlForceNone
                         'glb_TestInstance = theexec.DataManager.instancename
                         'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                 Next site
          Next i
                         Tname = CStr(Relay_Checked) & " On"
                         TheHdw.Utility.pins(Relay_Checked).State = tlUtilBitOn
                         '''TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
                         
                         For Each p In Pins_On
                                 'Tname = "rly_on_" & p
                                 BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)
                         
                                 'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
                                 TheExec.Datalog.WriteComment "Relay_on : " & BitState.pins(p)
                        Next p
          For i = 0 To UBound(PinArr_S)
                For Each site In TheExec.sites.Existing
                         TheHdw.Digital.pins(PinArr_S(i)).Disconnect
                         With TheHdw.PPMU.pins(PinArr_S(i))
                             .Connect
                             .ForceV FCurr, 0.04
                             .Gate = tlOn
                             
                         End With
                         
                 
                         TheHdw.Wait 0.001
                         'R_curr = thehdw.DCVS.Pins(rout_start).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                         R_Value = TheHdw.PPMU.pins(PinArr_S(i)).Read(tlPPMUReadMeasurements, 10)
                         'R_Value = R_Value.Math.Divide(thehdw.PPMU.Pins(PinArr_S(i)).current.Value)
                         R_Value = TheHdw.PPMU.pins(PinArr_S(i)).Voltage.value / TheHdw.PPMU.pins(PinArr_S(i)).Read(tlPPMUReadMeasurements, 10)
                         'Res_Value = R_curr / FCurr
                                 ''Res_Value = FVolt / R_curr
                         TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
                         'glb_TestInstance = theexec.DataManager.instancename
                         'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                         With TheHdw.PPMU.pins(PinArr_S(i))
                             .ForceV (0)
                             .Disconnect
                             .Gate = tlOff
                     
                         End With
                         
                   
                         
                           
                 Next site
        Next i
     Else
        For i = 0 To UBound(PinArr_D)
            For Each site In TheExec.sites.Existing
                
                         TheHdw.PPMU.pins(PinArr_D(i)).ClampVHi = 1#
                          
                          TheHdw.PPMU.pins(PinArr_D(i)).ForceV 0#
                          
                          Tname = CStr(Relay_Checked) & " On"
                          TheHdw.Utility.pins(Relay_Checked).State = tlUtilBitOn
                          '''TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
                          
                          For Each p In Pins_On
                                  'Tname = "rly_on_" & p
                                  BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)
                          
                                  'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
                                  TheExec.Datalog.WriteComment "Relay_on : " & BitState.pins(p)
                         Next p
                          
                          TheHdw.PPMU.pins(PinArr_D(i)).ForceI FCurr, FCurr
                     
                          TheHdw.Wait 0.002
                          R_Value = TheHdw.PPMU.pins(PinArr_D(i)).Read(tlPPMUReadMeasurements, 10)
                          
                          TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, hiVal:=Off_V_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="V", ForceResults:=tlForceNone
                          'glb_TestInstance = theexec.DataManager.instancename
                          'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                          With TheHdw.PPMU.pins(PinArr_D(i))
                              .ForceV (0)
                              .Disconnect
                              .Gate = tlOff
                      
                          End With
             Next site
        Next i
     End If
        TheHdw.Utility.pins(Relay_Checked).State = tlUtilBitOff
Exit Function


errHandler:
    ErrorDescription ("Relay_Resistance")
    If AbortTest Then Exit Function Else Resume Next
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''CAPACITR CHECK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function DCVS_MeasureCapacitor_2(sPinName As String, _
                                      sInstrumentType As String, _
                                      dbExpectedCapacitance As Double, _
                                      dbForceCurrent As Double, _
                                      dbForceVoltage As Double, Lo_Limit As Double, Hi_Limit As Double, _
                                      Optional dblParallelResistor As Double = -1, _
                                      Optional strParasiticRelayPin As String = vbNullString, _
                                      Optional blRelayON As Boolean = True) ',LowLimit As Double, Hilimit As Double)
    
    '********************capacitor main test begin********************'
    Dim dbCaptureTime As Double
    Dim dbIRange As Double
    Dim dSampleRate As Double
    Dim lCaptrueDepth As Long
    Dim dbMinForceCurrent As Double
    Dim dbMaxForceCurrent As Double
    Dim ilimit_onrange As Double
    Dim sdbCapValue As New SiteDouble
    Dim sdbCapValue_Parasitic As SiteDouble
    Dim pldMeasureValue As New PinListData
    
    On Error GoTo errHandler

    ' Start with nominal values
    ilimit_onrange = 0
    dbCaptureTime = 0.05 ' 50ms
    'dbForceVoltage = 2.5  ' 2.5V should be ok for all instruments
    ' Calculate current to Force I = C*dV/dt
    'dbForceCurrent = dbExpectedCapacitance * dbForceVoltage / dbCaptureTime
        
    dbMinForceCurrent = 0.00001 ' 10uA minimum current
    dbMaxForceCurrent = 0.002    ' 2mA max current
    If sInstrumentType = "HEXVS" Then
        dSampleRate = 25000000#
        lCaptrueDepth = 256000#
        dbIRange = 1
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.5
    ElseIf sInstrumentType = "HDVS" Then
        dSampleRate = 5555555#
        lCaptrueDepth = 2048#
        dbIRange = 1
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.1
    ElseIf sInstrumentType = "VHDVS" Then
       If dbExpectedCapacitance > 0.0001 Then
            dSampleRate = 100000#
       Else
            dSampleRate = 100000#
        End If
        'dSampleRate = 200000#
        lCaptrueDepth = 16384#
        dbIRange = 0.02
        ilimit_onrange = 0.125
    ElseIf sInstrumentType = "VSM" Then
        dSampleRate = 30000000#
        lCaptrueDepth = 256000#
        dbIRange = 1
    Else
        MsgBox "Unsupported DCVS type found : " & sInstrumentType
    End If
    
    ' If current is outside limits, set to limit and recalculate deltaT
    If dbForceCurrent < dbMinForceCurrent Then
        dbForceCurrent = dbMinForceCurrent
        dbCaptureTime = dbExpectedCapacitance * dbForceVoltage / dbForceCurrent
    ElseIf dbForceCurrent > dbMaxForceCurrent Then
        dbForceCurrent = dbMaxForceCurrent
        dbCaptureTime = dbExpectedCapacitance * dbForceVoltage / dbForceCurrent
    End If
    
    'If strParasiticRelayPin <> vbNullString Then
    '    If blRelayON = True Then
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOn
    '    Else
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOff
    '    End If
    '    Set sdbCapValue_Parasitic = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
    '
    '    If blRelayON = True Then
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOff
    '    Else
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOn
    '    End If
    '    Set sdbCapValue = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
    '    sdbCapValue = sdbCapValue_Parasitic.Subtract(sdbCapValue)
    'Else
      '''''  Set sdbCapValue = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
      '**************************************************************************************************************************************************************************************
      Dim lSampleNum As Long
    Dim sWaveSignalName As String
    Dim dspWave_VCapture As New DSPWave
    Dim dspWave_ICapture As New DSPWave
    Dim vLow As Double
    Dim vHigh As Double
    Dim lIndexlow As Long
    Dim lIndexHigh As Long
    Dim dbTimeDelta As Double
    Dim vsite As Variant
    'Dim sdbCapValue As New SiteDouble
            
    sWaveSignalName = "DIBC_DCVS_CapMeasureSig"
    vLow = 0.3 * dbForceVoltage
    vHigh = 0.7 * dbForceVoltage
    lSampleNum = CLng(dSampleRate * dbCaptureTime)
    If lSampleNum > lCaptrueDepth Then lSampleNum = lCaptrueDepth
    'TheExec.Datalog.WriteComment ("DCVS_CapMeasure with Pin " & sPinName & " and number of samples " & lSampleNum)
    
    ' Connect the DCVS pin forcing 0V and source current clamps at dbForceCurrent
    With TheHdw.DCVS.pins(sPinName)
        .Alarm(tlDCVSAlarmAll) = tlAlarmOff ' disable alarms
        .Voltage.Main.value = 0
        'Call SetDCVSCurrentAndRange(sPinName, dbForceCurrent / (1 + ilimit_onrange), dbIRange)
      '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                Dim savedAlarmSetting As tlAlarmBehavior
                Dim dbNewIRange As Double
                Dim rangeOptions() As Double
                Dim i As Long
                Dim rangeIncreasing As Boolean
                
                On Error GoTo errHandler
                If TheExec.sites.Selected.Count = 0 Then Exit Function
            
                dbNewIRange = dbIRange
                ' Get list of current range options
                rangeOptions = TheHdw.DCVS.pins(sPinName).CurrentRange.list
                
                If UBound(rangeOptions) >= 1 Then
                    ' Are the values in the list increasing?
                    If rangeOptions(0) < rangeOptions(1) Then
                        rangeIncreasing = True
                    Else
                        rangeIncreasing = False
                    End If
                    
                    ' Determine the new Range
                    For i = 0 To UBound(rangeOptions)
                        If rangeIncreasing Then 'range values are increasing
                            If dbForceCurrent <= rangeOptions(i) Then
                                    dbNewIRange = rangeOptions(i)
                                Exit For
                            End If
                        Else ' range values are decreasing
                            If dbForceCurrent > rangeOptions(i) Then
                                If i > 0 Then
                                    dbNewIRange = rangeOptions(i - 1)
                                Else
                                    dbNewIRange = rangeOptions(0)
                                End If
                                Exit For
                            End If
                        End If
                    Next i
                End If
                   
                ' Are we increasing current?
                If TheHdw.DCVS.pins(sPinName).CurrentLimit.Source.FoldLimit.level.value < dbForceCurrent Then
                    ' First increase the range, then the current
                    TheHdw.DCVS.pins(sPinName).CurrentRange = dbNewIRange
                    TheHdw.DCVS.pins(sPinName).Meter.CurrentRange = dbNewIRange
                    TheHdw.DCVS.pins(sPinName).CurrentLimit.Source.FoldLimit.level.value = dbForceCurrent
                Else  ' we are decreasing current
                  
                     TheHdw.DCVS.pins(sPinName).CurrentRange = dbNewIRange
                    TheHdw.DCVS.pins(sPinName).CurrentLimit.Source.FoldLimit.level.value = dbForceCurrent
                   
                    TheHdw.DCVS.pins(sPinName).Meter.CurrentRange = dbNewIRange
                End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '.CurrentRange = dbIRange
        '.CurrentLimit.Source.FoldLimit.level.value = dbForceCurrent / (1 + ilimit_onrange)
        .mode = tlDCVSModeVoltage
        .Meter.mode = tlDCVSMeterVoltage
        .Meter.CurrentRange = dbIRange
        .Connect (tlDCVSConnectDefault)
        .Gate = True
        Call TheHdw.Wait(0.2)
    End With
        
    ' Set up the capture signal
    With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterVoltage
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lSampleNum
        .Signals(sWaveSignalName).LoadSettings
    End With
    Call TheHdw.Wait(0.01)
      
    ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger
    
    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage
    
    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend
    
    ' Read the capture wave and process the results
    dspWave_VCapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave
    
    ' Go back to 0V
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = 0
    TheHdw.Wait (2 * dbCaptureTime)
    TheHdw.Wait (0.15)
    
    ' Set up current capture signal
    With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterCurrent
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lSampleNum
        .Signals(sWaveSignalName).LoadSettings
    End With
    
    Call TheHdw.Wait(0.001)
     ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger
    
    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage
    
    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend
    
    ' Read the capture wave and process the results
    dspWave_ICapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave
    
    For Each vsite In TheExec.sites.Selected
        lIndexlow = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vLow)
        lIndexHigh = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vHigh)
        sdbCapValue.value = -1
        If lIndexlow > 1 And lIndexHigh < dspWave_VCapture.SampleSize - 1 And lIndexlow < lIndexHigh Then
            dbTimeDelta = (lIndexHigh - lIndexlow) / dSampleRate
            dbForceCurrent = dspWave_ICapture.Select(lIndexlow, 1, lIndexHigh - lIndexlow).CalcMean
            If dblParallelResistor < 0 Then  'no parallel resistor
                'computing of the capacitance value: C= I * dt / dU
                sdbCapValue.value = dbForceCurrent * (dbTimeDelta / (vHigh - vLow))
            Else
                'computing of the capacitance value: C= dt * (I/dU - 1/R)
                sdbCapValue.value = dbTimeDelta * (dbForceCurrent / (vHigh - vLow) - 1 / dblParallelResistor)
            End If
        End If
    Next vsite
      '******************************************************************************************************************************************************************************************
    'End If
    Dim Tname As String
    Tname = sPinName & " Cap measure"
    pldMeasureValue.AddPin sPinName
    pldMeasureValue.pins(sPinName) = sdbCapValue
    TheExec.flow.TestLimit resultVal:=sdbCapValue, PinName:=sPinName, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=dbForceVoltage, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = 0
       TheHdw.Wait (0.15)
    TheHdw.DCVS.pins(sPinName).Disconnect
    TheHdw.DCVS.pins(sPinName).Gate = False
    
     TheHdw.Wait (0.15)
'''''    pldMeasureValue.AddPin sPinName
'''''    pldMeasureValue.Pins(sPinName) = sdbCapValue
'''''
'''''    Set DCVS_MeasureCapacitor = pldMeasureValue

TheHdw.DCVS.pins(sPinName).ClearCaptureMemory


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function UVI80_DPDT_Relay_check(R_LowLimit As Double, R_HiLimit As Double, FCurr As Double, Off_V_LowLimit As Double, Off_V_HiLimit As Double, Optional Relay_Checked As PinList, Optional WaitTime As Double = 0.003, Optional Default_pin As PinList, Optional Switch_pin As PinList)

'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim Wait_Time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    'Dim R_Volt As PinListData
    Dim R_Value As New PinListData
    Dim Res_Value As Double
    Dim FVolt As Double
    Dim i As Integer
    Dim PinArr_D() As String
    Dim PinCount_D As Long
    Dim PinArr_S() As String
    Dim PinCount_S As Long
    'Dim FCurr As Double
'''    Dim HiLimit As Double
'''    Dim LowLimit As Double
    On Error GoTo errHandler
''''    HiLimit = 300#
''''    LowLimit = 200#
    'FVolt = 0.02
    'FCurr = 0.0004
    
    TheHdw.DCVS.pins("All_power").Connect tlDCVSConnectDefault
                     TheHdw.Wait 0.01
                     With TheHdw.DCVS.pins("All_power")
                        .mode = tlDCVSModeVoltage
                        .Voltage.Main.value = 0#
                        .Voltage.Alt.value = 0#
                        .Voltage.Output = tlDCVSVoltageMain
                        .CurrentRange.value = 0.02
                        .Meter.mode = tlDCVSMeterCurrent
                        .Meter.CurrentRange = 0.02
                        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
                        .Gate = True
                        TheHdw.Wait 1#
                      End With
    
    
    
    
    relayOnStr = ""
    relayOffStr = ""
  
    TheExec.DataManager.DecomposePinList Relay_Checked, Pins_On(), Pin_Cnt_On

    
'''    Trim_NC_Pin Pins_On, Pin_Cnt_On
'''    Trim_NC_Pin Pins_Off, Pin_Cnt_Off
    
    'theexec.Datalog.WriteComment "============================ Relay ON ====================================="
    Tname = CStr(Relay_Checked) & " Off"
    TheHdw.Utility.pins(Relay_Checked).State = tlUtilBitOff
    For Each p In Pins_On
            'Tname = "rly_on_" & p
            BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)

        'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
        TheExec.Datalog.WriteComment "Relay_off : " & BitState.pins(p)
    Next p
   TheExec.DataManager.DecomposePinList Default_pin, PinArr_D, PinCount_D
   For i = 0 To UBound(PinArr_D)
           For Each site In TheExec.sites.Existing
                    TheHdw.DCVI.pins(PinArr_D(i)).Connect tlDCVIConnectDefault
                    TheHdw.DCVI.pins(PinArr_D(i)).Disconnect tlDCVIConnectHighSense
                    TheHdw.DCVI.pins(PinArr_D(i)).LocalKelvin = True
                    TheHdw.Wait 0.0005
                    '***-----------------------***'
            
                    '********************relay test begin********************'
                    
            
                    '***set DCVI instrument***'
                    With TheHdw.DCVI.pins(PinArr_D(i))
                                
                            '.mode = tlDCVIModeCurrent
                            .mode = tlDCVIModeVoltage
                            .ComplianceRange(tlDCVICompliancePositive).value = 7
                            .ComplianceRange(tlDCVIComplianceNegative).value = 2
                            TheHdw.Wait 0.01
                            .Current = 0.2
                            .Voltage = FCurr
                            .VoltageRange.value = 2
                            .CurrentRange.value = 0.2
                            .NominalBandwidth = 1003
                            '.Meter.mode = tlDCVIMeterVoltage
                            .Meter.mode = tlDCVIMeterCurrent
                            '.Meter.VoltageRange = 2
                            .Meter.CurrentRange = 0.2
                            .Meter.Filter.bypass = False
                            .Meter.Filter.value = 10000
                            .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                            .Gate = True
                            TheHdw.Wait 0.01
                    End With
                     'Dim pldMeasureValue As New PinListData
                     Dim UVI80_measurement As Double
                     Dim R_calc_Value As Double
                     'pldMeasureValue = thehdw.DCVI.Pins(PinArr_D(i)).Meter.Read(tlStrobe, 10, 100000, tlDCVIMeterReadingFormatAverage)
                     UVI80_measurement = TheHdw.DCVI.pins(PinArr_D(i)).Meter.Read(tlStrobe, 10, 100000, tlDCVIMeterReadingFormatAverage)
                  
                     R_calc_Value = TheHdw.DCVI.pins(PinArr_D(i)).Voltage / UVI80_measurement
                    TheExec.flow.TestLimit resultVal:=R_calc_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                    'glb_TestInstance = theexec.DataManager.instancename
                    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                    With TheHdw.DCVI.pins(PinArr_D(i))
                        .Voltage = 0
                        .Disconnect
                        .Gate = False
                
                    End With
           Next site
     Next i
       TheExec.DataManager.DecomposePinList Switch_pin, PinArr_S, PinCount_S
    For i = 0 To UBound(PinArr_S)
           For Each site In TheExec.sites.Existing
                    TheHdw.Digital.pins(PinArr_S(i)).Disconnect
                    TheHdw.PPMU.pins(PinArr_S(i)).ClampVHi = 1#
                      With TheHdw.PPMU.pins(PinArr_S(i))
                        .Connect
                        .ForceI 0.002, 0.002
                        .Gate = tlOn
                        
                    End With
                    'thehdw.PPMU.Pins(PinArr_S(i)).ForceV 0#
                    'thehdw.PPMU.Pins(PinArr_S(i)).ForceI 0.0002, 0.002
                   
                    R_Value = TheHdw.PPMU.pins(PinArr_S(i)).Read(tlPPMUReadMeasurements, 10)
                    TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, hiVal:=Off_V_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="V", ForceResults:=tlForceNone
                    'glb_TestInstance = theexec.DataManager.instancename
                    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            Next site
     Next i
                    Tname = CStr(Relay_Checked) & " On"
                    TheHdw.Utility.pins(Relay_Checked).State = tlUtilBitOn
                    '''TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
                    
                    For Each p In Pins_On
                            'Tname = "rly_on_" & p
                            BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)
                    
                            'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
                            TheExec.Datalog.WriteComment "Relay_on : " & BitState.pins(p)
                   Next p
     For i = 0 To UBound(PinArr_S)
           For Each site In TheExec.sites.Existing
                    TheHdw.Digital.pins(PinArr_S(i)).Disconnect
                    With TheHdw.PPMU.pins(PinArr_S(i))
                        .Connect
                        .ForceV FCurr, 0.04
                        .Gate = tlOn
                        
                    End With
                    
            
                    TheHdw.Wait 0.001
                    'R_curr = thehdw.DCVS.Pins(rout_start).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                    R_Value = TheHdw.PPMU.pins(PinArr_S(i)).Read(tlPPMUReadMeasurements, 10)
                    'R_Value = R_Value.Math.Divide(thehdw.PPMU.Pins(PinArr_S(i)).current.Value)
                    R_Value = TheHdw.PPMU.pins(PinArr_S(i)).Voltage.value / TheHdw.PPMU.pins(PinArr_S(i)).Read(tlPPMUReadMeasurements, 10)
                    'Res_Value = R_curr / FCurr
                            ''Res_Value = FVolt / R_curr
                    TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                    'glb_TestInstance = theexec.DataManager.instancename
                    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                    With TheHdw.PPMU.pins(PinArr_S(i))
                        .ForceV (0)
                        .Disconnect
                        .Gate = tlOff
                
                    End With
                    
              
                    
                      
            Next site
   Next i
   TheHdw.Utility.pins(Relay_Checked).State = tlUtilBitOff
Exit Function


errHandler:
    ErrorDescription ("Relay_Resistance")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Rly_Normal_Close_check(R_LowLimit As Double, R_HiLimit As Double, FCurr As Double, Off_V_LowLimit As Double, Off_V_HiLimit As Double, Optional Relay_On As PinList, Optional WaitTime As Double = 0.003, Optional Route_Start As PinList, Optional Route_End As PinList)

'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim Wait_Time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    'Dim R_Volt As PinListData
    Dim R_Value As New PinListData
    Dim Res_Value As Double
    Dim FVolt As Double
    'Dim FCurr As Double
'''    Dim HiLimit As Double
'''    Dim LowLimit As Double
    On Error GoTo errHandler
''''    HiLimit = 300#
''''    LowLimit = 200#
    'FVolt = 0.02
    'FCurr = 0.0004
    relayOnStr = ""
    relayOffStr = ""
  
    TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On

    
'''    Trim_NC_Pin Pins_On, Pin_Cnt_On
'''    Trim_NC_Pin Pins_Off, Pin_Cnt_Off
    
    'theexec.Datalog.WriteComment "============================ Relay ON ====================================="
    Tname = CStr(Relay_On) & " Off"
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
    For Each p In Pins_On
            'Tname = "rly_on_" & p
            BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)

        'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
        TheExec.Datalog.WriteComment "Relay_off : " & BitState.pins(p)
    Next p
   For Each site In TheExec.sites.Existing
        If LCase(TheExec.DataManager.pinType(Route_End)) = "power" Then
            With TheHdw.DCVS.pins(Route_End)
                .Connect
                .mode = tlDCVSModeVoltage
                .Voltage.Main = 0#
                .Gate = True

            End With
        ElseIf LCase(TheExec.DataManager.pinType(Route_End)) = "i/o" Then
            TheHdw.Digital.pins(Route_Start).Disconnect
            With TheHdw.PPMU.pins(Route_End)
                .Connect
                .ForceV 0#
                .Gate = tlOn

            End With
        ElseIf Route_End = "" Then

        End If
        
            TheHdw.Digital.pins(Route_Start).Disconnect
            With TheHdw.PPMU.pins(Route_Start)
                .Connect
                .ForceI FCurr, FCurr
                '.ForceV 0.1, FCurr
                .Gate = tlOn
                
            End With
        
             TheHdw.Wait 0.001
            'R_curr = thehdw.DCVS.Pins(rout_start).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
            R_Value = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
            R_Value = R_Value.Math.divide(TheHdw.PPMU.pins(Route_Start).Current.value)
            'Res_Value = thehdw.PPMU.Pins(Route_Start).Voltage / thehdw.PPMU.Pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
                    ''Res_Value = FVolt / R_curr
            TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=R_LowLimit, hiVal:=R_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
            'glb_TestInstance = theexec.DataManager.instancename
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            TheHdw.PPMU.pins(Route_Start).ClampVHi = 1#
            
            TheHdw.PPMU.pins(Route_Start).ForceV 0#
            
            Tname = CStr(Relay_On) & " On"
            TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
            '''TheExec.Datalog.WriteComment "Relay_off : " & BitState.Pins(p)
            
            For Each p In Pins_On
                    'Tname = "rly_on_" & p
                    BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)
            
                    'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
                    TheExec.Datalog.WriteComment "Relay_on : " & BitState.pins(p)
           Next p
            
            TheHdw.PPMU.pins(Route_Start).ForceI FCurr, FCurr
       '    thehdw.PPMU.Pins(Route_Start).ForceV 0.1, FCurr
            TheHdw.Wait 0.002
            R_Value = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
            
            TheExec.flow.TestLimit resultVal:=R_Value, lowVal:=Off_V_LowLimit, hiVal:=Off_V_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="V", ForceResults:=tlForceNone
            'glb_TestInstance = theexec.DataManager.instancename
            'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
            With TheHdw.PPMU.pins(Route_Start)
                .ForceV (0)
                .Disconnect
                .Gate = tlOff
        
            End With
            
       If LCase(TheExec.DataManager.pinType(Route_End)) = "power" Then
            With TheHdw.DCVS.pins(Route_End)
                .mode = tlDCVSModeVoltage
                .Voltage.Main = 0#
                .Disconnect
                .Gate = False
            End With
        ElseIf LCase(TheExec.DataManager.pinType(Route_End)) = "i/o" Then
            TheHdw.Digital.pins(Route_Start).Disconnect
            With TheHdw.PPMU.pins(Route_End)
                .ForceV 0#
                .Disconnect
                .Gate = tlOff
                
            End With
        ElseIf Route_End = "" Then
        
        End If
            
              TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
    Next site
Exit Function


errHandler:
    ErrorDescription ("Relay_Resistance")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UVI80_Relay_check_Short(R_LowLimit As Double, R_HiLimit As Double, FCurr As Double, Off_V_LowLimit As Double, Off_V_HiLimit As Double, Optional Relay_On As PinList, Optional WaitTime As Double = 0.003, Optional Route_Start As PinList, Optional Route_End As PinList)

'''''''control relay on off, will auto trim NC pins Tto cover CP, FT both stages
''''''    Dim Pins_On() As String, Pin_Cnt_On As Long
''''''    Dim p As Variant
''''''    Dim relayOnStr As String, relayOffStr As String
''''''    Dim Wait_Time As Double 'relay wiat time by global spec
''''''    Dim Tname As String
''''''    Dim BitState As New PinListData
''''''    'Dim R_Volt As PinListData
''''''    Dim R_Value As New PinListData
''''''    Dim Res_Value As Double
''''''    Dim FVolt As Double
''''''    'dim U_pins as String
''''''    On Error GoTo errhandler
''''''    'FVolt = 0.02
''''''    'FCurr = 0.0004
''''''    relayOnStr = ""
''''''    relayOffStr = ""
''''''
''''''    Theexec.DataManager.DecomposePinList relay_on, Pins_On(), Pin_Cnt_On
''''''
''''''
''''''    'theexec.Datalog.WriteComment "============================ Relay ON ====================================="
''''''    Tname = CStr(relay_on) & " On"
''''''   thehdw.Utility.Pins(relay_on).State = tlUtilBitOn
''''''    For Each p In Pins_On
''''''            'Tname = "rly_on_" & p
''''''            BitState = thehdw.Utility.Pins(p).States(tlUBStateProgrammed)
''''''
''''''        'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
''''''        Theexec.Datalog.WriteComment "Relay_on : " & BitState.Pins(p)
''''''    Next p
''''''    Dim Irange As Double
''''''    Irange = 0.02
''''''
''''''   For Each site In Theexec.sites.Existing
''''''        'thehdw.DCVI.Pins(Route_Start).Connect
''''''        With thehdw.DCVI.Pins(Route_Start)
''''''
'''''''''            .mode = tlDCVIModeVoltage
'''''''''            '.SetCurrentAndRange Irang, Irang
'''''''''            .VoltageRange = 1#
'''''''''            .CurrentRange.Value = Irange
'''''''''            '.current = Irange
'''''''''            '.Meter.CurrentRange = Irange
'''''''''            .Gate = True
''''''             .Gate = False
''''''            .Disconnect tlDCVIConnectDefault
''''''            .mode = tlDCVIModeHighImpedance
''''''            ' Connect only the sense to use with high impedance mode
''''''            .Connect tlDCVIConnectHighSense
''''''            .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
''''''            .current = 0
''''''            .Gate = True
''''''        End With
''''''          thehdw.DCVI.Pins(Route_Start).Voltage = FCurr
''''''          thehdw.Wait 0.001
''''''          R_Value = thehdw.DCVI.Pins(Route_Start).Meter.Read
''''''           thehdw.DCVI.Pins(Route_Start).Voltage = 0#
''''''           thehdw.DCVI.Pins(Route_Start).Gate = False
''''''           thehdw.DCVI.Pins(Route_Start).Disconnect
''''''
''''''
''''''
''''''            Theexec.Flow.TestLimit resultVal:=R_Value, lowval:=Off_V_LowLimit, hival:=Off_V_HiLimit, Unit:=unitCustom, ScaleType:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=FCurr, ForceUnit:=unitVolt, customUnit:="V", ForceResults:=tlForceNone
''''''
''''''
''''''    Next site
    
     Dim vsite As Variant
      
    
    For Each vsite In TheExec.sites.Active

        '***connect DCVI instrument***'
        TheHdw.DCVI.pins("HSC_VBG").Connect tlDCVIConnectDefault
        TheHdw.DCVI.pins("HSC_VBG").Disconnect tlDCVIConnectHighSense
        TheHdw.DCVI.pins("HSC_VBG").LocalKelvin = True
        TheHdw.Wait 0.0005
        '***-----------------------***'

        '********************relay test begin********************'
        TheHdw.Utility.pins("K23").State = tlUtilBitOff
        TheHdw.Wait 0.005

        '***set DCVI instrument***'
        With TheHdw.DCVI.pins("HSC_VBG")
                    
                .mode = tlDCVIModeCurrent
                '.mode = tlDCVIModeVoltage
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Current = 0.0015
                .Voltage = 1
                .VoltageRange.value = 2
                .CurrentRange.value = 0.002
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                '.Meter.mode = tlDCVIMeterCurrent
                .Meter.VoltageRange = 2
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                .Gate = True
                TheHdw.Wait 0.01
        End With
     
        '***-----------------------***'


        '*********************Relay main test begin********************'
        Dim pldMeasureValue As New PinListData

        Set pldMeasureValue = TheHdw.DCVI.pins("HSC_VBG").Meter.Read(tlStrobe, 10, 100000, tlDCVIMeterReadingFormatAverage)

       ' judgeCheckResult pldMeasureValue, 28084, "K98 ON"
       ' If Theexec.sites.ActiveCount = 0 Then Exit Sub

        '********************main test end********************'

        TheHdw.Utility.pins("K23").State = tlUtilBitOff
        TheHdw.Wait 0.003

        '********************subtest begin 1********************'
        pldMeasureValue = TheHdw.DCVI.pins("HSC_VBG").Meter.Read(tlStrobe, 10, 100000, tlDCVIMeterReadingFormatAverage)

        'judgeCheckResult pldMeasureValue, 28085, "K98 OFF"
        'If Theexec.sites.ActiveCount = 0 Then Exit Sub

        '********************subtest end 1********************'

        '********************relay test end********************'

        TheHdw.DCVI.pins("HSC_VBG").Alarm(tlDCVIAlarmMode) = tlAlarmDefault

        '***disconnect DCVI instrument***'
        With TheHdw.DCVI.pins("HSC_VBG")
                .Gate = False
                TheHdw.Wait 0.002
                .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 0.0005
        '***-----------------------***'

        '***disconnect DCTime instrument***'
       '' thehdw.DCTime.Pins("PAD_MTR_VREF_N_DCTIME").Disconnect
        '***-----------------------***'
    Next vsite
Exit Function


errHandler:
    ErrorDescription ("Relay_Check")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function K1_3P01KR_check(Optional Relay_On As PinList, Optional WaitTime As Double = 0.003, Optional Route_Start As PinList, Optional Route_End As PinList)

'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim Wait_Time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    Dim R_curr As Double
    Dim Res_Value As Double
    Dim FVolt As Double
    Dim FCurr As Double
    Dim hiLimit As Double
    Dim LowLimit As Double
    On Error GoTo errHandler
    hiLimit = 3500#
    LowLimit = 2500#
    FVolt = 0.1
    FCurr = 0.0002
    relayOnStr = ""
    relayOffStr = ""
    Tname = CStr(Relay_On) & "_3P01KR"
    TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On

    
'''    Trim_NC_Pin Pins_On, Pin_Cnt_On
'''    Trim_NC_Pin Pins_Off, Pin_Cnt_Off
    
    TheExec.Datalog.WriteComment "============================ Relay ON ====================================="
  
    TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
    For Each p In Pins_On
            'Tname = "rly_on_" & p
            BitState = TheHdw.Utility.pins(p).States(tlUBStateProgrammed)

        'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
        TheExec.Datalog.WriteComment "Relay_on : " & BitState.pins(p)
    Next p
   For Each site In TheExec.sites.Existing
'''''        With thehdw.DCVS.Pins(Route_End)
'''''            .mode = tlDCVSModeVoltage
'''''            .Voltage.Main = 0#
'''''            .Gate = True
'''''            .Connect
'''''        End With
        
        TheHdw.Digital.pins(Route_Start).Disconnect
        With TheHdw.PPMU.pins(Route_Start)
            .ForceI FCurr
            .Gate = tlOn
            .Connect
        End With
    
         TheHdw.Wait 0.001
        'R_curr = thehdw.DCVS.Pins(rout_start).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        R_curr = TheHdw.PPMU.pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
        Res_Value = R_curr / FCurr
        TheExec.flow.TestLimit resultVal:=Res_Value, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        With TheHdw.PPMU.pins(Route_Start)
            .ForceV (0)
            .Disconnect
            .Gate = tlOff
    
        End With
''''        With thehdw.DCVS.Pins(Route_End)
''''            .Voltage.Main = 0
''''            .Disconnect
''''            .Gate = False
''''        End With
    Next site
  
Exit Function
errHandler:
    ErrorDescription ("Relay_Control_1")
    If AbortTest Then Exit Function Else Resume Next
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''CAPACITR CHECK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function DCVS_MeasureCapacitor(sPinName As String, _
                                      sInstrumentType As String, _
                                      dbExpectedCapacitance As Double, _
                                      dbForceCurrent As Double, _
                                      dbForceVoltage As Double, LowLimit As Double, hiLimit As Double, _
                                      Optional dblParallelResistor As Double = -1, _
                                      Optional strParasiticRelayPin As String = vbNullString, _
                                      Optional blRelayON As Boolean = True) As Long 'As PinListData
    
    '********************capacitor main test begin********************'
    Dim dbCaptureTime As Double
    Dim dbIRange As Double
    Dim IRange_F As Double
    Dim dSampleRate As Double
    Dim lCaptrueDepth As Long
    Dim dbMinForceCurrent As Double
    Dim dbMaxForceCurrent As Double
    Dim ilimit_onrange As Double
    'Dim sdbCapValue As SiteDouble
    Dim sdbCapValue_Parasitic As SiteDouble
    Dim pldMeasureValue As New PinListData
    Dim sWaveSignalName As String
    Dim Tname As String
    Dim vLow As Double
    Dim vHigh As Double
    Dim dspWave_VCapture As New DSPWave
    Dim dspWave_ICapture As New DSPWave
    Dim vsite As Variant
    Dim lIndexlow As Long
    Dim lIndexHigh As Long
    Dim dbTimeDelta As Double
    Dim sdbCapValue As New SiteDouble
    ''''Dim LowLimit As Double
    ''''Dim HiLimit As Double
    On Error GoTo errHandler
   
''''     LowLimit = 0.000821476
''''     HiLimit = 0.00141804
    ' Start with nominal values
    ilimit_onrange = 0
    dbCaptureTime = 0.05 ' 50ms
    vLow = 0.3 * dbForceVoltage
    vHigh = 0.7 * dbForceVoltage
     sWaveSignalName = "DIBC_DCVS_CapMeasureSig"
    'dbForceVoltage = 2.5  ' 2.5V should be ok for all instruments
    ' Calculate current to Force I = C*dV/dt
    'dbForceCurrent = dbExpectedCapacitance * dbForceVoltage / dbCaptureTime
        
    dbMinForceCurrent = 0.00001 ' 10uA minimum current
    dbMaxForceCurrent = 0.002    ' 2mA max current
    If sInstrumentType = "HEXVS" Then
       ' dSampleRate = 25000000#
             dSampleRate = 1562500#
        lCaptrueDepth = 256000#
        dbIRange = 1
        IRange_F = 15
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.5
    ElseIf sInstrumentType = "HDVS" Then
        dSampleRate = 5555555#
        lCaptrueDepth = 2048#
        dbIRange = 1
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.1
    ElseIf sInstrumentType = "VHDVS" Then
        If dbExpectedCapacitance > 0.0001 Then
            dSampleRate = 100000#
       Else
            dSampleRate = 200000#
        End If
        lCaptrueDepth = 16384#
        IRange_F = 0.02
        dbIRange = 0.02
        dbMinForceCurrent = 0.02
        ilimit_onrange = 0.125
    ElseIf sInstrumentType = "VSM" Then
        dSampleRate = 30000000#
        lCaptrueDepth = 256000#
        dbIRange = 1
    Else
        MsgBox "Unsupported DCVS type found : " & sInstrumentType
    End If
    
    ' If current is outside limits, set to limit and recalculate deltaT
''''    If dbForceCurrent < dbMinForceCurrent Then
''''        dbForceCurrent = dbMinForceCurrent
''''        dbCaptureTime = dbExpectedCapacitance * dbForceVoltage / dbForceCurrent
''''    ElseIf dbForceCurrent > dbMaxForceCurrent Then
''''        dbForceCurrent = dbMaxForceCurrent
''''        dbCaptureTime = dbExpectedCapacitance * dbForceVoltage / dbForceCurrent
''''    End If
''''
    
    With TheHdw.DCVS.pins(sPinName)
        .Alarm(tlDCVSAlarmAll) = tlAlarmOff ' disable alarms
        .Voltage.Main.value = 0
       ' Call SetDCVSCurrentAndRange(sPinName, dbForceCurrent / (1 + ilimit_onrange), dbIRange)
        .CurrentRange = IRange_F
        .CurrentLimit.Source.FoldLimit.level.value = dbMinForceCurrent 'dbForceCurrent / (1 + ilimit_onrange)
        .mode = tlDCVSModeVoltage
        .Meter.mode = tlDCVSMeterVoltage
        .Meter.CurrentRange = dbIRange
        .Connect (tlDCVSConnectDefault)
        .Gate = True
    End With
        Call TheHdw.Wait(0.1)
    
        With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterVoltage
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lCaptrueDepth
        .Signals(sWaveSignalName).LoadSettings
    End With
    Call TheHdw.Wait(0.1)
      
    ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger
    
    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage
    
    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend
    
    ' Read the capture wave and process the results
    dspWave_VCapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave
'''    Dim Voltage_Scope As String
'''   Voltage_Scope="Profile V"
'''    dspWave_VCapture.Plot (Voltage_Scope)
    ' Go back to 0V
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = 0
    'thehdw.Wait (2 * dbCaptureTime)
    TheHdw.Wait (0.1)
    
    ' Set up current capture signal
    With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterCurrent
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lCaptrueDepth
        .Signals(sWaveSignalName).LoadSettings
    End With
    
    Call TheHdw.Wait(0.001)
     ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger
    
    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage
    
    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend

    ' Read the capture wave and process the results
    dspWave_ICapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave
''''    Dim Current_Scope As String
''''     Current_Scope = "Profile I"
''''    dspWave_ICapture.Plot (Current_Scope)

    For Each vsite In TheExec.sites.Selected
        lIndexlow = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vLow)
        lIndexHigh = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vHigh)
        sdbCapValue.value = -1
        'If lIndexlow > 1 And lIndexHigh < dspWave_VCapture.SampleSize - 1 And lIndexlow < lIndexHigh Then
            dbTimeDelta = (lIndexHigh - lIndexlow) / dSampleRate
            dbForceCurrent = dspWave_ICapture.Select(lIndexlow, 1, lIndexHigh - lIndexlow).CalcMean
            'dspwave_icapture.Select
            If dblParallelResistor < 0 Then  'no parallel resistor
                'computing of the capacitance value: C= I * dt / dU
                sdbCapValue.value = dbForceCurrent * (dbTimeDelta / (vHigh - vLow))
            Else
                'computing of the capacitance value: C= dt * (I/dU - 1/R)
                sdbCapValue.value = dbTimeDelta * (dbForceCurrent / (vHigh - vLow) - 1 / dblParallelResistor)
            End If
       ' End If
    Next vsite
    
    'Set DCVS_MeasureCapacitorOnce = sdbCapValue
    
    
    
    
    
    'If strParasiticRelayPin <> vbNullString Then
    '    If blRelayON = True Then
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOn
    '    Else
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOff
    '    End If
    '    Set sdbCapValue_Parasitic = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
    '
    '    If blRelayON = True Then
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOff
    '    Else
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOn
    '    End If
    '    Set sdbCapValue = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
    '    sdbCapValue = sdbCapValue_Parasitic.Subtract(sdbCapValue)
    'Else
   '     Set sdbCapValue = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
    'End If
    Tname = sPinName & " Cap measure"
    pldMeasureValue.AddPin sPinName
    pldMeasureValue.pins(sPinName) = sdbCapValue
    TheExec.flow.TestLimit resultVal:=sdbCapValue, PinName:=sPinName, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=dbForceVoltage, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    'Set DCVS_MeasureCapacitor = pldMeasureValue
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''CAPACITR CHECK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function DCVS_MeasureCapacitor_check(sPinName As String, _
                                      sInstrumentType As String, _
                                      dbExpectedCapacitance As Double, _
                                      dbForceVoltage As Double, LowLimit As Double, hiLimit As Double, _
                                      Optional dblParallelResistor As Double = -1, _
                                      Optional strParasiticRelayPin As String = vbNullString, _
                                      Optional blRelayON As Boolean = True) As Long 'As PinListData
    
    '********************capacitor main test begin********************'
    Dim dbCaptureTime As Double
    Dim dbIRange As Double
    Dim IRange_F As Double
    Dim dSampleRate As Double
    Dim lCaptrueDepth As Long
    Dim dbMinForceCurrent As Double
    Dim dbMaxForceCurrent As Double
    Dim ilimit_onrange As Double
    'Dim sdbCapValue As SiteDouble
    Dim sdbCapValue_Parasitic As SiteDouble
    Dim pldMeasureValue As New PinListData
    Dim sWaveSignalName As String
    Dim Tname As String
    Dim vLow As Double
    Dim vHigh As Double
    Dim dspWave_VCapture As New DSPWave
    Dim dspWave_ICapture As New DSPWave
    Dim vsite As Variant
    Dim lIndexlow As Long
    Dim lIndexHigh As Long
    Dim dbTimeDelta As Double
    Dim sdbCapValue As New SiteDouble
    Dim dbForceCurrent As Double
    ''''Dim LowLimit As Double
    ''''Dim HiLimit As Double
    On Error GoTo errHandler
   
''''     LowLimit = 0.000821476
''''     HiLimit = 0.00141804
    ' Start with nominal values
    ilimit_onrange = 0
    dbCaptureTime = 0.05 ' 50ms
    vLow = 0.3 * dbForceVoltage
    vHigh = 0.7 * dbForceVoltage
     sWaveSignalName = "DIBC_DCVS_CapMeasureSig"
    'dbForceVoltage = 2.5  ' 2.5V should be ok for all instruments
    ' Calculate current to Force I = C*dV/dt
    'dbForceCurrent = dbExpectedCapacitance * dbForceVoltage / dbCaptureTime
        
    dbMinForceCurrent = 0.00001 ' 10uA minimum current
    dbMaxForceCurrent = 0.002    ' 2mA max current
    If sInstrumentType = "HEXVS" Then
        dSampleRate = 5000000#
       ' dSampleRate = 1562500#
        lCaptrueDepth = 256000#
        dbIRange = 1
        IRange_F = 15
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.5
    ElseIf sInstrumentType = "HDVS" Then
        dSampleRate = 5555555#
        lCaptrueDepth = 2048#
        dbIRange = 1
        dbMinForceCurrent = 0.05
        dbMaxForceCurrent = 0.1
    ElseIf sInstrumentType = "VHDVS" Then
        If dbExpectedCapacitance > 0.0001 Then
            dSampleRate = 100000#
       Else
            dSampleRate = 200000#
        End If
        lCaptrueDepth = 16384#
        IRange_F = 0.02
        dbIRange = 0.02
        dbMinForceCurrent = 0.02
        ilimit_onrange = 0.125
    ElseIf sInstrumentType = "VSM" Then
        dSampleRate = 30000000#
        lCaptrueDepth = 256000#
        dbIRange = 1
    Else
        MsgBox "Unsupported DCVS type found : " & sInstrumentType
    End If
    
    ' If current is outside limits, set to limit and recalculate deltaT
''''    If dbForceCurrent < dbMinForceCurrent Then
''''        dbForceCurrent = dbMinForceCurrent
''''        dbCaptureTime = dbExpectedCapacitance * dbForceVoltage / dbForceCurrent
''''    ElseIf dbForceCurrent > dbMaxForceCurrent Then
''''        dbForceCurrent = dbMaxForceCurrent
''''        dbCaptureTime = dbExpectedCapacitance * dbForceVoltage / dbForceCurrent
''''    End If
''''
    
    With TheHdw.DCVS.pins(sPinName)
        .Disconnect
        .Alarm(tlDCVSAlarmAll) = tlAlarmOff ' disable alarms
        .Voltage.Main.value = 0
       ' Call SetDCVSCurrentAndRange(sPinName, dbForceCurrent / (1 + ilimit_onrange), dbIRange)
        .CurrentRange = IRange_F
        .CurrentLimit.Source.FoldLimit.level.value = dbMinForceCurrent 'dbForceCurrent / (1 + ilimit_onrange)
        .mode = tlDCVSModeVoltage
        .Meter.mode = tlDCVSMeterVoltage
        .Meter.CurrentRange = dbIRange
        .Connect (tlDCVSConnectDefault)
        .Gate = True
    End With
        Call TheHdw.Wait(0.1)
    
        With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterVoltage
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lCaptrueDepth
        .Signals(sWaveSignalName).LoadSettings
    End With
    Call TheHdw.Wait(0.05)
      
    ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger
    
    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage
    
    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend
    
    ' Read the capture wave and process the results
    dspWave_VCapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave
''''    Dim Voltage_Scope As String
''''   Voltage_Scope = "Profile V"
''''    dspWave_VCapture.Plot (Voltage_Scope)
    ' Go back to 0V
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = 0
    'thehdw.Wait (2 * dbCaptureTime)
    TheHdw.Wait (0.1)
    
    ' Set up current capture signal
    With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterCurrent
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lCaptrueDepth
        .Signals(sWaveSignalName).LoadSettings
    End With
    
    Call TheHdw.Wait(0.001)
     ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger
    
    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage
    
    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend

    ' Read the capture wave and process the results
    dspWave_ICapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave
''''    Dim Current_Scope As String
''''     Current_Scope = "Profile I"
''''    dspWave_ICapture.Plot (Current_Scope)
    Dim Vmax_cap As Double
    Dim Vmin_cap As Double
    For Each vsite In TheExec.sites.Selected
        lIndexlow = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vLow)
        lIndexHigh = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vHigh)
        Vmin_cap = dspWave_VCapture.Element(lIndexlow)
        Vmax_cap = dspWave_VCapture.Element(lIndexHigh)
        sdbCapValue.value = -1
        'If lIndexlow > 1 And lIndexHigh < dspWave_VCapture.SampleSize - 1 And lIndexlow < lIndexHigh Then
            dbTimeDelta = (lIndexHigh - lIndexlow) / dSampleRate
            dbForceCurrent = dspWave_ICapture.Select(lIndexlow, 1, lIndexHigh - lIndexlow).CalcMean
            'dspwave_icapture.Select
            If dblParallelResistor < 0 Then  'no parallel resistor
                'computing of the capacitance value: C= I * dt / dU
               ' sdbCapValue.Value = dbForceCurrent * (dbTimeDelta / (vHigh - vLow))
                sdbCapValue.value = dbForceCurrent * (dbTimeDelta / (Vmax_cap - Vmin_cap))
            Else
                'computing of the capacitance value: C= dt * (I/dU - 1/R)
                sdbCapValue.value = dbTimeDelta * (dbForceCurrent / (vHigh - vLow) - 1 / dblParallelResistor)
            End If
       ' End If
    Next vsite
    
    'Set DCVS_MeasureCapacitorOnce = sdbCapValue
    
    
    
    
    
    'If strParasiticRelayPin <> vbNullString Then
    '    If blRelayON = True Then
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOn
    '    Else
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOff
    '    End If
    '    Set sdbCapValue_Parasitic = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
    '
    '    If blRelayON = True Then
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOff
    '    Else
    '        TheHdw.Utility.Pins(strParasiticRelayPin).State = tlUtilBitOn
    '    End If
    '    Set sdbCapValue = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
    '    sdbCapValue = sdbCapValue_Parasitic.Subtract(sdbCapValue)
    'Else
   '     Set sdbCapValue = DCVS_MeasureCapacitorOnce(sPinName, dbForceCurrent, dbForceVoltage, dSampleRate, lCaptrueDepth, dbCaptureTime, dbIRange, ilimit_onrange, dblParallelResistor)
    'End If
    Tname = sPinName & " Cap measure"
    pldMeasureValue.AddPin sPinName
    pldMeasureValue.pins(sPinName) = sdbCapValue
    TheExec.flow.TestLimit resultVal:=sdbCapValue, PinName:=sPinName, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=dbForceVoltage, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    With TheHdw.DCVS.pins(sPinName)
        .CurrentRange = 0.02
        .Voltage.Main.value = 0#
        .Gate = False
       TheHdw.Wait (0.002)
        .Disconnect (tlDCVSConnectDefault)
 
    End With
    'Set DCVS_MeasureCapacitor = pldMeasureValue
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Private Function DCVS_MeasureCapacitorOnce(sPinName As String, dbForceCurrent As Double, dbForceVoltage As Double, dSampleRate As Double, lCaptrueDepth As Long, _
                                             dbCaptureTime As Double, dbIRange As Double, Optional ilimit_onrange As Double = 0, _
                                             Optional dblParallelResistor As Double = -1) As SiteDouble
     
    Dim lSampleNum As Long
    Dim sWaveSignalName As String
    Dim dspWave_VCapture As New DSPWave
    Dim dspWave_ICapture As New DSPWave
    Dim vLow As Double
    Dim vHigh As Double
    Dim lIndexlow As Long
    Dim lIndexHigh As Long
    Dim dbTimeDelta As Double
    Dim vsite As Variant
    Dim sdbCapValue As New SiteDouble
            
    sWaveSignalName = "DIBC_DCVS_CapMeasureSig"
    vLow = 0.3 * dbForceVoltage
    vHigh = 0.7 * dbForceVoltage
    lSampleNum = CLng(dSampleRate * dbCaptureTime)
    If lSampleNum > lCaptrueDepth Then lSampleNum = lCaptrueDepth
    'TheExec.Datalog.WriteComment ("DCVS_CapMeasure with Pin " & sPinName & " and number of samples " & lSampleNum)
    
    ' Connect the DCVS pin forcing 0V and source current clamps at dbForceCurrent
''''    With thehdw.DCVS.Pins(sPinName)
''''        .Alarm(tlDCVSAlarmAll) = tlAlarmOff ' disable alarms
''''        .Voltage.Main.Value = 0
''''        Call SetDCVSCurrentAndRange(sPinName, dbForceCurrent / (1 + ilimit_onrange), dbIRange)
''''        '.CurrentRange = dbIRange
''''        '.CurrentLimit.Source.FoldLimit.level.value = dbForceCurrent / (1 + ilimit_onrange)
''''        .mode = tlDCVSModeVoltage
''''        .Meter.mode = tlDCVSMeterVoltage
''''        .Meter.CurrentRange = dbIRange
''''        .Connect (tlDCVSConnectDefault)
''''        .Gate = True
''''    End With
        Call TheHdw.Wait(0.1)
    ' Set up the capture signal
    With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterVoltage
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lSampleNum
        .Signals(sWaveSignalName).LoadSettings
    End With
    Call TheHdw.Wait(0.1)
      
    ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger
    
    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage
    
    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend
    
    ' Read the capture wave and process the results
    dspWave_VCapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave
'''    Dim Voltage_Scope As String
'''   Voltage_Scope="Profile V"
'''    dspWave_VCapture.Plot (Voltage_Scope)
    ' Go back to 0V
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = 0
    TheHdw.Wait (2 * dbCaptureTime)
    TheHdw.Wait (0.1)
    
    ' Set up current capture signal
    With TheHdw.DCVS.pins(sPinName).Capture
        .Signals.Add (sWaveSignalName)
        .Signals(sWaveSignalName).mode = tlDCVSMeterCurrent
        .Signals(sWaveSignalName).SampleRate = dSampleRate
        .Signals(sWaveSignalName).SampleSize = lSampleNum
        .Signals(sWaveSignalName).LoadSettings
    End With
    
    Call TheHdw.Wait(0.001)
     ' Trigger the capture
    Call TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).Trigger
    
    ' Force the new voltage
    TheHdw.DCVS.pins(sPinName).Voltage.Main.value = dbForceVoltage
    
    While TheHdw.DCVS.pins(sPinName).Capture.IsRunning
        TheHdw.Wait (0.001)
    Wend

    ' Read the capture wave and process the results
    dspWave_ICapture = TheHdw.DCVS.pins(sPinName).Capture.Signals(sWaveSignalName).DSPWave
    Dim Current_Scope As String
     Current_Scope = "Profile I"
    dspWave_ICapture.Plot (Current_Scope)

    For Each vsite In TheExec.sites.Selected
        lIndexlow = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vLow)
        lIndexHigh = dspWave_VCapture.FindIndex(OfLastElement, LessThan, vHigh)
        sdbCapValue.value = -1
        If lIndexlow > 1 And lIndexHigh < dspWave_VCapture.SampleSize - 1 And lIndexlow < lIndexHigh Then
            dbTimeDelta = (lIndexHigh - lIndexlow) / dSampleRate
            dbForceCurrent = dspWave_ICapture.Select(lIndexlow, 1, lIndexHigh - lIndexlow).CalcMean
            'dspwave_icapture.Select
            If dblParallelResistor < 0 Then  'no parallel resistor
                'computing of the capacitance value: C= I * dt / dU
                sdbCapValue.value = dbForceCurrent * (dbTimeDelta / (vHigh - vLow))
            Else
                'computing of the capacitance value: C= dt * (I/dU - 1/R)
                sdbCapValue.value = dbTimeDelta * (dbForceCurrent / (vHigh - vLow) - 1 / dblParallelResistor)
            End If
        End If
    Next vsite
    
    Set DCVS_MeasureCapacitorOnce = sdbCapValue
End Function
' This subroutine sets the current, range, and meter range for a dcvs
' It is intelligent about picking the lowest possible range
' It is also intelligent about the order to do things depending if current is increasing or decreasing
Private Sub SetDCVSCurrentAndRange(PinName As String, dbForceCurrent As Double, dbDefIRange As Double)

    Dim savedAlarmSetting As tlAlarmBehavior
    Dim dbNewIRange As Double
    Dim rangeOptions() As Double
    Dim i As Long
    Dim rangeIncreasing As Boolean
    
    On Error GoTo errHandler
    If TheExec.sites.Selected.Count = 0 Then Exit Sub

    dbNewIRange = dbDefIRange
    ' Get list of current range options
    rangeOptions = TheHdw.DCVS.pins(PinName).CurrentRange.list
    
    If UBound(rangeOptions) >= 1 Then
        ' Are the values in the list increasing?
        If rangeOptions(0) < rangeOptions(1) Then
            rangeIncreasing = True
        Else
            rangeIncreasing = False
        End If
        
        ' Determine the new Range
        For i = 0 To UBound(rangeOptions)
            If rangeIncreasing Then 'range values are increasing
                If dbForceCurrent <= rangeOptions(i) Then
                        dbNewIRange = rangeOptions(i)
                    Exit For
                End If
            Else ' range values are decreasing
                If dbForceCurrent > rangeOptions(i) Then
                    If i > 0 Then
                        dbNewIRange = rangeOptions(i - 1)
                    Else
                        dbNewIRange = rangeOptions(0)
                    End If
                    Exit For
                End If
            End If
        Next i
    End If
       
    ' Are we increasing current?
    If TheHdw.DCVS.pins(PinName).CurrentLimit.Source.FoldLimit.level.value < dbForceCurrent Then
        ' First increase the range, then the current
        TheHdw.DCVS.pins(PinName).CurrentRange = dbNewIRange
        TheHdw.DCVS.pins(PinName).Meter.CurrentRange = dbNewIRange
        TheHdw.DCVS.pins(PinName).CurrentLimit.Source.FoldLimit.level.value = dbForceCurrent
    Else  ' we are decreasing current
      
         TheHdw.DCVS.pins(PinName).CurrentRange = dbNewIRange
        TheHdw.DCVS.pins(PinName).CurrentLimit.Source.FoldLimit.level.value = dbForceCurrent
       
        TheHdw.DCVS.pins(PinName).Meter.CurrentRange = dbNewIRange
    End If

    Exit Sub
errHandler:
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function MEM0_cap(Optional Relay_On As PinList, Optional WaitTime As Double = 0.003, Optional Route_Start As PinList, Optional Route_End As PinList)



'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim Wait_Time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    Dim R_curr As Double
    Dim Res_Value As Double
    Dim FVolt As Double
    Dim FCurr As Double
    Dim hiLimit As Double
    Dim LowLimit As Double
    On Error GoTo errHandler
    hiLimit = 300#
    LowLimit = 200#
    FVolt = 0.02
    FCurr = 0.0004
    relayOnStr = ""
    relayOffStr = ""
    Tname = "profile"
''''    Theexec.DataManager.DecomposePinList relay_on, Pins_On(), Pin_Cnt_On

    
'''    Trim_NC_Pin Pins_On, Pin_Cnt_On
'''    Trim_NC_Pin Pins_Off, Pin_Cnt_Off
    
''''    Theexec.Datalog.WriteComment "============================ Relay ON ====================================="
''''
''''    thehdw.Utility.Pins(relay_on).state = tlUtilBitOn
''''    For Each p In Pins_On
''''            'Tname = "rly_on_" & p
''''            BitState = thehdw.Utility.Pins(p).States(tlUBStateProgrammed)
''''
''''        'TheExec.Flow.TestLimit resultval:=BitState.Pins(p), lowVal:=tlUtilBitOn, hiVal:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
''''        Theexec.Datalog.WriteComment "Relay_on : " & BitState.Pins(p)
''''    Next p
'''''   For Each site In Theexec.sites.Existing
        With TheHdw.DCVS.pins(Route_End)
            .mode = tlDCVSModeVoltage
            .SetCurrentRanges 0.2, 0.2
            .Voltage.Main = 0
            .Gate = True
            .Connect
            .Voltage.Main = 0.9
        End With
'''''
'''''        thehdw.Digital.Pins(Route_Start).Disconnect
'''''        With thehdw.PPMU.Pins(Route_Start)
'''''            .ForceV FVolt
'''''            .Gate = tlOn
'''''            .Connect
'''''        End With
    
         TheHdw.Wait 0.08
'''''        'R_curr = thehdw.DCVS.Pins(rout_start).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
'''''        R_curr = thehdw.PPMU.Pins(Route_Start).Read(tlPPMUReadMeasurements, 10)
'''''        'Res_Value = R_curr / FCurr
'''''                Res_Value = FVolt / R_curr
'''''        Theexec.Flow.TestLimit resultVal:=Res_Value, lowval:=LowLimit, hival:=HiLimit, Unit:=unitCustom, ScaleType:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=FCurr, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
    
'''        With thehdw.PPMU.Pins(Route_Start)
'''            .ForceV (0)
'''            .Disconnect
'''            .Gate = tlOff
'''
'''        End With
        With TheHdw.DCVS.pins(Route_End)
            .Voltage.Main = 0
            .Disconnect
            .Gate = False
        End With
          'thehdw.Utility.Pins(relay_on).state = tlUtilBitOff
    'next site
Exit Function


errHandler:
    ErrorDescription ("Relay_Resistance")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DIB_ReadTDR_Value() As Long
      
    Dim strAllDigitalPins() As String
    Dim strAllDigitalPins_ch() As String
    Dim sAllDigitalPins As String
    Dim sAllDigitalPins_ch As String
    Dim lngPnum As Long
    Dim lngChnum As Long
    Dim lngSnum As Long
    Dim strError As String
    Dim Non_ExcludedCal_IO() As String
    
    Dim LowLimit As Double    '0.0000000005
    Dim hiLimit As Double
    Dim Tname As String
    LowLimit = 0.0000000005
    hiLimit = 0.000000005
    Tname = "IO_TDR_Value"
    Call TheExec.DataManager.GetPinNames(strAllDigitalPins, chIO, lngPnum)
    sAllDigitalPins = Join(strAllDigitalPins, ",")

    Dim blnReturnSignalNames As Boolean
    blnReturnSignalNames = TheExec.DataManager.ReturnSignalNames
    
    TheExec.DataManager.ReturnSignalNames = True
    Call TheExec.DataManager.GetPinNames(strAllDigitalPins, chIO, lngPnum)
    sAllDigitalPins = Join(strAllDigitalPins, ",")
    
''''    'skip the channels connected to clock buffer and PA channels to DUT
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "REFCLK_XI0,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "REFCLK_RT_CLK32768,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "FRC_BACKUP_1,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "FRC_BACKUP_4,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "FRC_BACKUP_5,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "FRC_BACKUP_2,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "REFCLK_MONITORCHANNEL,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "FRC_BACKUP_3,", "")
''''    'skip the PA channels connected to DUT
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "XI0_PA,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "RT_CLK32768_PA,", "")
''''
''''     'VDD_PCPU_MONITOR,VDD_GPU_MONITOR,VDD_SOC_MONITOR" conenct to cap->gnd
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_PCPU_MONITOR,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_GPU_MONITOR,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_SOC_MONITOR,", "")
''''    'excluded channels connect to caps or inductors
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX0_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX0_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX1_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX1_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX2_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX2_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX3_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX3_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX4_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RX4_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX0_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX0_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX1_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX1_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX2_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX2_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX3_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX3_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX4_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_TX4_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D0_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D0_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D1_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D1_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D2_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D2_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D3_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D3_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D4_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D4_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D5_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D5_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D6_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D6_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D7_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D7_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D8_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D8_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D9_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D9_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D10_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D10_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D11_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX_RX_D11_N,", "")
''''    'excluded channels connect to resistors
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_DCS_SENSE,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_DCS_SENSE_MONITOR,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_GPU_SENSE,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_GPU_SENSE_MONITOR,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_PCPU_SENSE,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_PCPU_SENSE_MONITOR,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_SOC_SENSE,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDD_SOC_SENSE_MONITOR,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDDQL_SENSE,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VDDQL_SENSE_MONITOR,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VSS_DDR_SENSE,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VSS_GPU_SENSE,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "VSS_PCPU_SENSE,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "GPIO23,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RCAL_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PCIE_RCAL_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "MIPI0C_REXT,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "MIPID_REXT,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDP_RCAL_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDP_RCAL_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX0_RCAL_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX0_RCAL_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX1_RCAL_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "LPDPRX1_RCAL_N,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DBG_USB_REXT,", "") 'could not adjust the sequency between DBG_USB_REXT and USB_REXT. Will cause wrong delete.
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "USB_REXT,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DBG_USB_DP,", "") 'could not adjust the sequency between DBG_USB_DP and USB_DP. Will cause wrong delete.
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DBG_USB_DM,", "") 'could not adjust the sequency between DBG_USB_DM and USB_DM. Will cause wrong delete.
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "USB_DP,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "USB_DM,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DDR0_ZQ,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DDR3_ZQ,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DDR0_RREF,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DDR1_RREF,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DDR2_RREF,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "DDR3_RREF,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PAD_MTR_ANALOG_TEST_P_TC,", "") 'Masked due to teradyne no test
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PAD_MTR_ANALOG_TEST_N_TC,", "") 'Masked due to teradyne no test
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PAD_MTR_VREF_P_TC,", "")        'Masked due to teradyne no test
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PAD_MTR_VREF_N_TC,", "")        'Masked due to teradyne no test
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PAD_MTR_RREF_P,", "")
''''    sAllDigitalPins = Replace(UCase(sAllDigitalPins), "PAD_MTR_RREF_N,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "VSS_1,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "VSS_1_MONITOR,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "VSS_2,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "VSS_2_MONITOR,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "VSS_3,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "VSS_3_MONITOR,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "VSS_4,", "")
'''''    sAllDigitalPins = Replace(Ucase(sAllDigitalPins), "VSS_4_MONITOR,", "")

    Non_ExcludedCal_IO = Split(sAllDigitalPins, ",")
    'Call TheExec.DataManager.GetChannelList(sAllDigitalPins, TL_NONUMBER, chIO, strAllDigitalPins_ch, lngChnum, lngSnum, strError)
    Dim sPin As Variant
    Dim sPin_ch() As String
    Dim lngTraceLength As Double
    Dim plTraceLength As New PinListData
    Dim thissite As Variant
    Dim siteNO As Integer
    
    For Each sPin In Non_ExcludedCal_IO
        plTraceLength.AddPin (sPin)
        'Call TheExec.DataManager.GetChannelList(sPin, TL_NONUMBER, chIO, sPin_ch, lngChnum, lngSnum, strError)
        'lngTraceLength = TheHdw.Digital.Calibration.Channels(sPin_ch).DIB.Trace
        Call TheExec.DataManager.GetChannelListForSelectedSites(sPin, chIO, sPin_ch, lngChnum, lngSnum, strError)
        siteNO = 0
        For Each thissite In TheExec.sites.Selected
            lngTraceLength = TheHdw.Digital.Calibration.Channels(sPin_ch(siteNO)).DIB.trace
            plTraceLength.pins(sPin).value = lngTraceLength
            siteNO = siteNO + 1
            If UBound(sPin_ch) = 0 Then siteNO = 0
        Next thissite
    Next sPin
          
        'lngTraceLength = TheHdw.Digital.Calibration.Channels(sPin).DIB.Trace
        'Judge Check Result
        'judgeCheckResult pldMeasureValue_, 2400, "Check_IO Connectivity"
        'TheExec.Datalog.WriteComment (sPin & Str(lngTraceLength))
        
    TheExec.DataManager.ReturnSignalNames = blnReturnSignalNames 'set back to default
    TheExec.flow.TestLimit resultVal:=plTraceLength, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="S", ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    'Set DIBC_MeasureTDR_New = plTraceLength
       
'    '''' just for debug usage
'    Dim lngTraceLength_1 As Double
'    lngTraceLength_1 = TheHdw.Digital.Calibration.Channels("2.ch0").DIB.Trace
End Function

Public Function Check_User_Power() As Long
    Dim PowerList() As String
    Dim PowerValue As Variant
    Dim ResultsValue As Double
    Dim loLimit As Double, hiLimit As Double
    Dim vForce As Double
    Dim vsite As Variant 'add by rita on 2018May23
    Dim pldMeasureValue As New PinListData 'add by rita on 2018May23
    Dim Tname As String
    Tname = "Check_UserPower"
    PowerList = TheHdw.DIB.power.list
    TheHdw.DIB.powerOn = False
    TheHdw.Wait 0.5
    
    For Each PowerValue In PowerList
        If PowerValue = "48V" Then
            vForce = 48
            loLimit = vForce * 0.95: hiLimit = vForce * 1.05
            TheHdw.DIB.power(PowerValue).State = tlOn
            TheHdw.Wait 0.1
            ResultsValue = TheHdw.DIB.power.item(PowerValue).Reading
                      
          pldMeasureValue.AddPin PowerValue  'add by rita on 2018May23
          For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(PowerValue).value(vsite) = ResultsValue
          Next vsite
          'judgeCheckResult pldMeasureValue, T_Num, "Check_UserPower", LoLimit, Hilimit, "V" 'add by rita on 2018May23
          TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=loLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="V", ForceResults:=tlForceNone
          'glb_TestInstance = theexec.DataManager.instancename
          'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
          TheHdw.DIB.power(PowerValue).State = tlOff
            
        ElseIf PowerValue = "15V_2" Then
            vForce = 15
            loLimit = vForce * 0.95: hiLimit = vForce * 1.05
            TheHdw.DIB.power(PowerValue).State = tlOn
            TheHdw.Wait 0.1
            ResultsValue = TheHdw.DIB.power.item(PowerValue).Reading
            
          'pldMeasureValue.AddPin PowerValue
          pldMeasureValue.pins.item(0).name = PowerValue 'edit by rita on 2018May23
          For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(PowerValue).value(vsite) = ResultsValue
          Next vsite
          'judgeCheckResult pldMeasureValue, T_Num + 1, "Check_UserPower", LoLimit, Hilimit, "V" 'add by rita on 2018May23
          TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=loLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="V", ForceResults:=tlForceNone
          'glb_TestInstance = theexec.DataManager.instancename
          'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        ElseIf PowerValue = "15V_1" Then
            vForce = 15
            loLimit = vForce * 0.95: hiLimit = vForce * 1.25
            TheHdw.DIB.power(PowerValue).State = tlOn
            TheHdw.Wait 0.1
            ResultsValue = TheHdw.DIB.power.item(PowerValue).Reading
            
          'pldMeasureValue.AddPin PowerValue
          pldMeasureValue.pins.item(0).name = PowerValue 'edit by rita on 2018May23
          For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(PowerValue).value(vsite) = ResultsValue
          Next vsite
          'judgeCheckResult pldMeasureValue, T_Num + 2, "Check_UserPower", LoLimit, Hilimit, "V" 'add by rita on 2018May23
          TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=loLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="V", ForceResults:=tlForceNone
          'glb_TestInstance = theexec.DataManager.instancename
          'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        ElseIf PowerValue = "12V" Then
            vForce = 12
            loLimit = vForce * 0.95: hiLimit = vForce * 1.25
            TheHdw.DIB.power(PowerValue).State = tlOn
            TheHdw.Wait 0.1
            ResultsValue = TheHdw.DIB.power.item(PowerValue).Reading
            
          'pldMeasureValue.AddPin PowerValue
          pldMeasureValue.pins.item(0).name = PowerValue 'edit by rita on 2018May23
          For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(PowerValue).value(vsite) = ResultsValue
          Next vsite
          'judgeCheckResult pldMeasureValue, T_Num + 3, "Check_UserPower", LoLimit, Hilimit, "V" 'comment by rita on 2018May23
          TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=loLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="V", ForceResults:=tlForceNone
          'glb_TestInstance = theexec.DataManager.instancename
          'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        ElseIf PowerValue = "5V_2" Then
            vForce = 5
            loLimit = vForce * 0.95: hiLimit = vForce * 1.05
            TheHdw.DIB.power(PowerValue).State = tlOn
            TheHdw.Wait 0.1
            ResultsValue = TheHdw.DIB.power.item(PowerValue).Reading
            
          'pldMeasureValue.AddPin PowerValue
          pldMeasureValue.pins.item(0).name = PowerValue 'edit by rita on 2018May23
          For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(PowerValue).value(vsite) = ResultsValue
          Next vsite
          'judgeCheckResult pldMeasureValue, T_Num + 4, "Check_UserPower", LoLimit, Hilimit, "V" 'add by rita on 2018May23
         TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=loLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="V", ForceResults:=tlForceNone
         'glb_TestInstance = theexec.DataManager.instancename
         'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        ElseIf PowerValue = "5V_1" Then
            vForce = 5
            loLimit = vForce * 0.95: hiLimit = vForce * 1.25
            TheHdw.DIB.power(PowerValue).State = tlOn
            TheHdw.Wait 0.1
            ResultsValue = TheHdw.DIB.power.item(PowerValue).Reading
                     
          'pldMeasureValue.AddPin PowerValue
          pldMeasureValue.pins.item(0).name = PowerValue 'edit by rita on 2018May23
          For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(PowerValue).value(vsite) = ResultsValue
          Next vsite
          'judgeCheckResult pldMeasureValue, T_Num + 5, "Check_UserPower", LoLimit, Hilimit, "V" 'add by rita on 2018May23
         TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=loLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="V", ForceResults:=tlForceNone
         'glb_TestInstance = theexec.DataManager.instancename
         'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        ElseIf PowerValue = "3.3V" Then
            vForce = 3.3
            loLimit = vForce * 0.95: hiLimit = vForce * 1.25
            TheHdw.DIB.power(PowerValue).State = tlOn
            TheHdw.Wait 0.1
            ResultsValue = TheHdw.DIB.power.item(PowerValue).Reading
            
          'pldMeasureValue.AddPin PowerValue
          pldMeasureValue.pins.item(0).name = PowerValue 'edit by rita on 2018May23
          For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(PowerValue).value(vsite) = ResultsValue
          Next vsite
          'judgeCheckResult pldMeasureValue, T_Num + 6, "Check_UserPower", LoLimit, Hilimit, "V" 'edit by rita on 2018May23
          TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=loLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="V", ForceResults:=tlForceNone
          'glb_TestInstance = theexec.DataManager.instancename
          'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        End If
    
    Next PowerValue

End Function

Public Function DIBCheckerTestFunction_CheckUDBState() As Long
                                                                                         
    Dim strAllRelayPinsArray()  As String
    Dim strAllRelayPinList As String
    Dim lngPnum As Long
                                                                                         
    Call TheExec.DataManager.GetPinNames(strAllRelayPinsArray, chUtil, lngPnum)
    strAllRelayPinList = Join(strAllRelayPinsArray, ",")
                                                                                         
    Dim pinnamex
    Dim pinx As String
    Dim loopk As Long
    Dim Tname As String
    
    pinnamex = Split(strAllRelayPinList, ",")
                                                                                         
    'relay on check
    Dim pldMeasureValue As New PinListData
                                                                   
    For loopk = 0 To UBound(pinnamex)
        Tname = "UDBCHECK_ON"
        pinx = pinnamex(loopk)
        Set pldMeasureValue = CheckUdbState(pinx, 0)
        TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=0, hiVal:=0, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        Set pldMeasureValue = Nothing
        'Call DIBCheckerTestFunction_CheckUDBState_ON(Tnum_ON + loopk, pinx)
    Next loopk
                                                                                         
    'relay off check
    For loopk = 0 To UBound(pinnamex)
        Tname = "UDBCHECK_OFF"
        pinx = pinnamex(loopk)
        Set pldMeasureValue = CheckUdbState(pinx, 1)
        TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=1, hiVal:=1, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                    'Call DIBCheckerTestFunction_CheckUDBState_OFF(Tnum_OFF + loopk, pinx)
        Set pldMeasureValue = Nothing
    Next loopk
                                                                                         
End Function

Public Function CheckUdbState(strRelayPinName As String, _
                                lngSetUdbState As Long) As PinListData
                               
    Dim pldRelayStateRes As New PinListData
   
    TheHdw.Utility.Threshold = 4
   
    'turn ON/ OFF relay and reab back its state
    If lngSetUdbState = 0 Then
        TheHdw.Utility.pins(strRelayPinName).State = tlUtilBitOn
    Else
        TheHdw.Utility.pins(strRelayPinName).State = tlUtilBitOff
    End If
    TheHdw.Wait 0.003
    'read back relay state for all sites
     Set CheckUdbState = TheHdw.Utility.pins(strRelayPinName).States(tlUBStateCompared)
   
    If TheExec.TesterMode = testModeOffline And lngSetUdbState = 0 Then CheckUdbState = CheckUdbState.Math.Abs.Add(1)
    
End Function

Public Function Read_Device_ID(Validating_ As Boolean)

    If Validating_ Then
        Call PrLoadPattern(".\Pattern\DIB\SPIROMReadID_Debug.PAT")
        Exit Function    ' Exit after validation
    End If
    

     TheHdw.Utility.pins("K02").State = tlUtilBitOn
     TheHdw.Utility.pins("K04").State = tlUtilBitOn
  '***connect DCVS instrument***'
   TheHdw.DCVS.pins("SPI_ROM_PWR").Connect tlDCVSConnectDefault
   TheHdw.Wait 0.0005
   '***-----------------------***'
   '***set DCVS instrument***'
   With TheHdw.DCVS.pins("SPI_ROM_PWR")
      .mode = tlDCVSModeVoltage
      TheHdw.Wait 0.01
      .Voltage.Main.value = 1.2 'AL20181114, 1.8v will cause the test result not stable.
      .Voltage.Alt.value = 1.2
      .Voltage.Output = tlDCVSVoltageMain
      .CurrentRange.value = 0.1
      .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff  'tlDCVSCurrentLimitBehaviorGateOff
      .CurrentLimit.Source.FoldLimit.level.value = 0.1
      .CurrentLimit.Source.FoldLimit.TimeOut.value = 0.05
      .Meter.mode = tlDCVSMeterVoltage
      .Meter.VoltageRange = 18
      .Meter.Filter.bypass = False
      .Alarm(tlDCVSAlarmAll) = tlAlarmDefault
      .Gate = True
      TheHdw.Wait 0.01
   End With
   '***-----------------------***'

    
'end of added by rita

Dim i As Integer
Dim pat_modify(7) As String
Dim numcap As Long
Dim PinData As New PinListData
Dim regdata() As Long
Dim rawdata As String
Dim modify_data As String
modify_data = ""
rawdata = ""
TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

TheHdw.Digital.pins("SPI0_MISO").Levels.DriverMode = tlDriverModeHiZ
TheHdw.PinLevels.pins("SPI0_MISO").ModifyLevel chVil, -0.15 '0828 by rita
TheHdw.PinLevels.pins("SPI0_MISO").ModifyLevel chVih, 1.8   '0828 by rita
TheHdw.PinLevels.pins("SPI0_MISO").ModifyLevel chVol, 0.8   '0828 by rita
TheHdw.PinLevels.pins("SPI0_MISO").ModifyLevel chVoh, 1    '0828 by rita
TheHdw.PinLevels.pins("SPI0_MISO").ModifyLevel chIol, 0.001    '0828 by rita
TheHdw.PinLevels.pins("SPI0_MISO").ModifyLevel chIoh, -0.001    '0828 by rita
TheHdw.PinLevels.pins("SPI0_MISO").ModifyLevel chVt, 0.9    '0828 by rita
TheHdw.Wait 0.05
Call SetupHRam
'TheHdw.Digital.Patterns.pat(".\Pattern\DIB\SPIROMReadID_Debug.PAT").Load
get_status_reg &H9F, 8  'Read ID


End Function


Public Function dib_nwire_test(T_Num As Long)

        '******************** Clock Fan-out Buffer ********************'
        'thehdw.DIB.powerOn = True
        'thehdw.Wait 0.005
        
        
        Call FreeRunClk_Disable_dibchecker("XI0_diff_Port") 'disable clock
        Call FreeRunClk_Disable_dibchecker("RT_CLK32768_Port") 'disable clock
        
        'Call subnwire("XI0_Diff_Port", "XI0_PA", "REFCLK_XI0", T_Num, "U1_REFCLK1", T_Num + 1, "U1_PLLLock1")
        Dim PortName As String:: PortName = "XI0_Diff_Port"
        Dim DUTPin As String:: DUTPin = "XI0_PA"
        Dim Bufferpin As String:: Bufferpin = "REFCLK_XI0"
        
        'Call subnwire("XI0_Diff_Port", "XI0_PA", "REFCLK_XI0", T_Num, "U1_REFCLK1", T_Num + 1, "U1_PLLLock1")
                    TheHdw.Digital.pins(PortName).Levels.value(chVil) = 0  'm_sAllDigitalPinList
                    TheHdw.Digital.pins(PortName).Levels.value(chVih) = 1.8
                    TheHdw.Digital.pins(PortName).Levels.value(chIol) = 0
                    TheHdw.Digital.pins(PortName).Levels.value(chIoh) = 0
                    TheHdw.Digital.pins(PortName).Levels.value(chVch) = 6
                    TheHdw.Digital.pins(PortName).Levels.value(chVoutLoTyp) = 0
                    TheHdw.Digital.pins(PortName).Levels.value(chVoutHiTyp) = 0
                    TheHdw.Digital.pins(PortName).Levels.DriverMode = tlDriverModeLargeHiZ  'Largeswing-HiZ
        
                    TheHdw.Digital.pins(DUTPin).Levels.value(chVol) = 0.54  'dut pin
                    TheHdw.Digital.pins(DUTPin).Levels.value(chVoh) = 0.54
                    TheHdw.Digital.pins(DUTPin).Levels.value(chVt) = 0.9
                    TheHdw.Digital.pins(DUTPin).Levels.value(chVcl) = 0
                    
                              
                    TheHdw.Digital.pins(Bufferpin).Levels.value(chVol) = 0.55
                    TheHdw.Digital.pins(Bufferpin).Levels.value(chVoh) = 0.55
                    
                    TheHdw.Digital.pins(Bufferpin).Levels.value(chVt) = 0.6
                    TheHdw.Digital.pins(Bufferpin).Levels.value(chVcl) = -1
                    TheHdw.Wait 0.01
                    '''
                    
                    TheHdw.Digital.pins(Bufferpin).Connect
                    TheHdw.Wait 0.005
                    '***************************************************
                    'Call StartSupportBoardClock_dibchecker
                       Dim SBFreq As Double
                       Dim SB_Vih As Double
                       SB_Vih = 1.8
            
                      'SBFreq = TheExec.Specs.Globals("FREQ_SB_VAR").ContextValue
                       SBFreq = 40 * MHz '24 * MHz '24000000# ''24MHZ
            
            
                       With TheHdw.DIB.SupportBoardClock
                          .Connect
                          .Frequency = SBFreq
                          .Vih = SB_Vih ' Max is 6V
                          .Vil = 0 ' Min is -1V
                          .start
                       End With
                    '********************************************************************
                    '-------------------------------------------------------------------------
                    MeasFreq_dibchecker Bufferpin, 0.001, "REFCLK_XIO" ', Tnum_5000, Tname_5000 ', True
                    '----------------------------------------------------------------------------
                    Call FreeRunclk_Enable_dibchecker(PortName, DUTPin, Bufferpin, "REFCLK_XIO_PLLLock") ', Tnum_5001, Tname_5001) 'Port Name
                    
                    TheHdw.Digital.pins(Bufferpin).Disconnect
        Call FreeRunClk_Disable_dibchecker("XI0_diff_Port") 'disable clock
        Call FreeRunClk_Disable_dibchecker("RT_CLK32768_Port") 'disable clock
        
        '==================================================================================================================
        'Call subnwire("RT_CLK32768_Port", "RT_CLK32768_PA", "REFCLK_RT_CLK32768", T_Num + 2, "U1_REFCLK2", T_Num + 3, "U1_PLLLock2")
         PortName = "RT_CLK32768_Port"
         DUTPin = "RT_CLK32768_PA"
         Bufferpin = "REFCLK_RT_CLK32768"
         TheHdw.Digital.pins(PortName).Levels.value(chVil) = 0  'm_sAllDigitalPinList
                    TheHdw.Digital.pins(PortName).Levels.value(chVih) = 1.8
                    TheHdw.Digital.pins(PortName).Levels.value(chIol) = 0
                    TheHdw.Digital.pins(PortName).Levels.value(chIoh) = 0
                    TheHdw.Digital.pins(PortName).Levels.value(chVch) = 6
                    TheHdw.Digital.pins(PortName).Levels.value(chVoutLoTyp) = 0
                    TheHdw.Digital.pins(PortName).Levels.value(chVoutHiTyp) = 0
                    TheHdw.Digital.pins(PortName).Levels.DriverMode = tlDriverModeLargeHiZ  'Largeswing-HiZ
        
                    TheHdw.Digital.pins(DUTPin).Levels.value(chVol) = 0.54  'dut pin
                    TheHdw.Digital.pins(DUTPin).Levels.value(chVoh) = 0.54
                    TheHdw.Digital.pins(DUTPin).Levels.value(chVt) = 0.9
                    TheHdw.Digital.pins(DUTPin).Levels.value(chVcl) = 0
                    
                              
                    TheHdw.Digital.pins(Bufferpin).Levels.value(chVol) = 0.55
                    TheHdw.Digital.pins(Bufferpin).Levels.value(chVoh) = 0.55
                    
                    TheHdw.Digital.pins(Bufferpin).Levels.value(chVt) = 0.6
                    TheHdw.Digital.pins(Bufferpin).Levels.value(chVcl) = -1
                    TheHdw.Wait 0.01
                    '''
                    
                    TheHdw.Digital.pins(Bufferpin).Connect
                    TheHdw.Wait 0.005
                    Call StartSupportBoardClock_dibchecker
                    
                    MeasFreq_dibchecker Bufferpin, 0.001, "REFCLK_RT_CLK32768" ' , Tnum_5000, Tname_5000 ', True
                    
                    Call FreeRunclk_Enable_dibchecker(PortName, DUTPin, Bufferpin, "REFCLK_RT_CLK32768_PLLLock") ', Tnum_5001, Tname_5001) 'Port Name
                    
                    TheHdw.Digital.pins(Bufferpin).Disconnect
        '==================================================================================================================
        Call FreeRunClk_Disable_dibchecker("XI0_diff_Port") 'disable clock
        Call FreeRunClk_Disable_dibchecker("RT_CLK32768_Port") 'disable clock

End Function

Public Function FreeRunClk_Disable_dibchecker(PortName As String) As Long
    Dim site As Variant
    
    TheHdw.Protocol.ports(PortName).Halt
    TheHdw.Protocol.ports(PortName).Enabled = False     'scope out point

End Function


Public Function StartSupportBoardClock_dibchecker() As Long
    On Error GoTo errHandler
    Dim SBFreq As Double
    Dim SB_Vih As Double
    SB_Vih = 1.8
    
    'SBFreq = TheExec.Specs.Globals("FREQ_SB_VAR").ContextValue
    SBFreq = 40 * MHz '24 * MHz '24000000# ''24MHZ
    
    
    With TheHdw.DIB.SupportBoardClock
        .Connect
        .Frequency = SBFreq
        .Vih = SB_Vih ' Max is 6V
        .Vil = 0 ' Min is -1V
        .start
    End With
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MeasFreq_dibchecker(mea_pin As String, period As Double, Tname As String, Optional collectdata As Boolean = False) As Long ', Tnum As Long, Tname As String, Optional collectdata As Boolean = False) As Long

    Dim freqCnt1 As New PinListData
    Dim MeasFreq1 As New SiteDouble
    Dim TimeInterval1 As New SiteDouble
    Dim StartVolt As Double
    Dim stepvolt As Double
    Dim stepcount As Integer
    Dim PinNum As Integer
    Dim i As Integer
    Dim currvolt As Double
    Dim site As Variant
    
    StartVolt = 0.1
    stepvolt = 0.05
    stepcount = 30
    If (collectdata = True) Then
        For i = 0 To stepcount
        
        currvolt = StartVolt + i * stepvolt
    
            TheHdw.Digital.pins(mea_pin).Levels.value(chVoh) = currvolt
            With TheHdw.Digital.pins(mea_pin).FreqCtr
                .EventSource = VOH
                .EventSlope = Positive
                .Enable = IntervalEnable
                .Interval = period
            End With
            
            TheHdw.Digital.pins(mea_pin).FreqCtr.Clear
            TheHdw.Digital.pins(mea_pin).FreqCtr.start
            freqCnt1 = TheHdw.Digital.pins(mea_pin).FreqCtr.Read
            TimeInterval1 = TheHdw.Digital.pins(mea_pin).FreqCtr.Interval
            MeasFreq1 = freqCnt1.Math.divide(TimeInterval1)
            Debug.Print ("measPin: " + mea_pin + " Voh=" + CStr(TheHdw.Digital.pins(mea_pin).Levels.value(chVoh)) + " Frequcney=" + CStr(MeasFreq1(0) / 1000000#) + " MHz")
            TheExec.Datalog.WriteComment ("measPin: " + mea_pin + " Voh=" + CStr(TheHdw.Digital.pins(mea_pin).Levels.value(chVoh)) + " Frequcney=" + CStr(MeasFreq1(0) / 1000000#) + " MHz")
      
                 
        Next i
    Else
        With TheHdw.Digital.pins(mea_pin).FreqCtr
            .EventSource = VOH
            .EventSlope = Positive
            .Enable = IntervalEnable
            .Interval = period
        End With
        
        TheHdw.Digital.pins(mea_pin).FreqCtr.Clear
        TheHdw.Digital.pins(mea_pin).FreqCtr.start
        freqCnt1 = TheHdw.Digital.pins(mea_pin).FreqCtr.Read
       For PinNum = 0 To freqCnt1.pins.Count - 1
            TimeInterval1 = TheHdw.Digital.pins(freqCnt1.pins(PinNum).name).FreqCtr.Interval
      
        
            MeasFreq1 = freqCnt1.pins(PinNum).divide(TimeInterval1)
            For Each site In TheExec.sites.Active
'              theexec.Datalog.WriteComment ("Site: " + CStr(site) + "  measPin: " + freqCnt1.Pins(pinnum).Name + " Frequcney=" + CStr(MeasFreq1(site) / 1000000#) + " MHz")
'              Debug.Print ("Site: " + CStr(site) + "  measPin: " + freqCnt1.Pins(pinnum).Name + " Frequcney=" + CStr(MeasFreq1(site) / 1000000#) + " MHz")
                                       
            Next site
            
          
            Dim pldMeasureValue As New PinListData
            pldMeasureValue.AddPin freqCnt1.pins(PinNum).name
            pldMeasureValue.pins(freqCnt1.pins(PinNum).name) = MeasFreq1
            
           ' judgeCheckResult pldMeasureValue, Tnum, Tname
           TheExec.flow.TestLimit resultVal:=pldMeasureValue, PinName:=PinNum, lowVal:=39996000, hiVal:=40000800, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="Hz", ForceResults:=tlForceNone
           'glb_TestInstance = theexec.DataManager.instancename
           'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
             
             
        Next PinNum
        
        
  ''      MeasFreq = MeasFreq1
    
    End If
   
End Function


Public Function FreeRunclk_Enable_dibchecker(PortName As String, DUTPin As String, Bufferpin As String, Tname As String) ', Tnum As Long, Tname As String, Optional ApplyLT As Boolean = True)

    Dim site As Variant
    Dim i As Long
    Dim PLLLockChecked As New SiteLong
    Dim measf As New PinListData
    Dim NotLocked As Boolean
    Dim XI0_REFCLK As String
    Dim PortMode As String
    Dim PLL_Lock As New SiteLong
    Dim FreeRunPortList() As String
    Dim PortNum As Integer
    Dim PrintInfo As Boolean
    
    On Error GoTo errHandler
    
    PrintInfo = True 'False
    
    FreeRunPortList = Split(PortName, ",")
    For Each site In TheExec.sites.Active
      PLLLockChecked = 1
      PLL_Lock = 1
    Next site
    
    TheHdw.Protocol.ports(PortName).Halt
    TheHdw.Protocol.ports(PortName).Enabled = False
    'End If

'        If TheExec.DataManager.instanceName <> "" Then
'            TheHdw.Digital.ApplyLevelsTiming True, ApplyLT, ApplyLT, tlPowered
'        Else
'            TheHdw.Digital.ApplyLevelsTiming False, True, True, tlPowered, , , , "LVL_nWire", "Normal", "NV", "TIMESET_nWire", "Normal", "typ"
'        End If
        
    TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, PinLevelsSheet:="DIBC_TSets__", TimeSetSheet:="DIBC_Levels__" ''May2 added
        '''0831 by rita
        TheHdw.Digital.pins(PortName).Levels.value(chVil) = 0  'm_sAllDigitalPinList
        TheHdw.Digital.pins(PortName).Levels.value(chVih) = 1.8
        TheHdw.Digital.pins(PortName).Levels.value(chIol) = 0
        TheHdw.Digital.pins(PortName).Levels.value(chIoh) = 0
        TheHdw.Digital.pins(PortName).Levels.value(chVch) = 6
        TheHdw.Digital.pins(PortName).Levels.value(chVoutLoTyp) = 0
        TheHdw.Digital.pins(PortName).Levels.value(chVoutHiTyp) = 0
        TheHdw.Digital.pins(PortName).Levels.DriverMode = tlDriverModeLargeHiZ  'Largeswing-HiZ
        
        TheHdw.Digital.pins(DUTPin).Levels.value(chVol) = 0.54
        TheHdw.Digital.pins(DUTPin).Levels.value(chVoh) = 0.54
        TheHdw.Digital.pins(DUTPin).Levels.value(chVt) = 0.9
        TheHdw.Digital.pins(DUTPin).Levels.value(chVcl) = 0
        
        TheHdw.Digital.pins(Bufferpin).Levels.value(chVol) = 0.55
        TheHdw.Digital.pins(Bufferpin).Levels.value(chVoh) = 0.55
        TheHdw.Digital.pins(Bufferpin).Levels.value(chVt) = 0.6
        TheHdw.Digital.pins(Bufferpin).Levels.value(chVcl) = -1
            ''''
        TheHdw.Wait 0.005 '0904 by rita
        TheHdw.Protocol.ports(PortName).Enabled = True
        TheHdw.Protocol.ports(PortName).NWire.ResetPLL
        TheHdw.Wait 0.001
        ' Test the PLL lock status and datalog results.
        For Each site In TheExec.sites.Active
          For PortNum = 0 To UBound(FreeRunPortList)
           If TheHdw.Protocol.ports(FreeRunPortList(PortNum)).NWire.IsPLLLocked = False Then
             If (PrintInfo = True) Then
                 'TheExec.Datalog.WriteComment ("******************  Site " + CStr(site) + " " + FreeRunPortList(PortNum) + " is not Locked ******************  ")
             End If
             PLL_Lock = 0
           Else
             If (PrintInfo = True) Then
                'TheExec.Datalog.WriteComment ("******************  Site " + CStr(site) + " " + FreeRunPortList(PortNum) + " is Locked" + " ******************  ")
             End If
           End If
          Next
        Next site
        
        ' Start the nWire engine.
        Call TheHdw.Protocol.ports(PortName).NWire.Frames("RunFreeClock").Execute
        TheHdw.Protocol.ports(PortName).IdleWait
        TheHdw.Wait 0.001
        
        ''''' measure freq
'        Call MeasFreq_dibchecker(DUTPin, 0.000001, True)     '''Apr28 added RT_CLK32768_PA
'        Call MeasFreq_dibchecker(Bufferpin, 0.000001, True) '''Apr28 added  REFCLK_RT_CLK32768
        'Call MeasFreq_dibchecker("RT_CLK32768_PA", 0.000001, True)     '''Apr28 added
       '' Call MeasFreq_dibchecker("REFCLK_RT_CLK32768", 0.000001, True) '''Apr28 added
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites.Selected
                PLL_Lock(site) = 1
            Next site
        End If
        
        Dim pldMeasureValue As New PinListData
        pldMeasureValue.AddPin PortName
        pldMeasureValue.pins(PortName) = PLL_Lock
        
       ' judgeCheckResult pldMeasureValue, Tnum, Tname
        TheExec.flow.TestLimit resultVal:=pldMeasureValue, PinName:=PortName, lowVal:=1, hiVal:=1, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")


    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function get_status_reg(data As Integer, data_len As Integer) As String
Dim i As Integer
Dim pat_modify() As String
Dim numcap As Long
Dim PinData As New PinListData
Dim regdata() As Long
Dim rawdata As String
Dim modify_data As String
modify_data = ""
get_status_reg = ""
ReDim pat_modify(data_len - 1)
Call get_data(data, pat_modify, data_len)
For i = 0 To data_len - 1
modify_data = modify_data + pat_modify(i)
Next i
Call TheHdw.Digital.pins("SPI1_MOSI").patterns(".\Pattern\DIB\SPIROMReadID_Debug.PAT").ModifyVectorBlockData("", 4, pat_modify)
TheHdw.Wait 0.05
'thehdw.Digital.Patterns.Pat(".\Pattern\DIB\IB21_SPIROMReadID_Debug.PAT").Load
';thehdw.Patterns(".\Pattern\DIB\IB21_PIROMReadID_Debug.PAT").Load
TheHdw.patterns(".\Pattern\DIB\SPIROMReadID_Debug.PAT").start
TheHdw.Digital.Patgen.HaltWait
numcap = 160  'update by tyrone on 3.24
'numcap = thehdw.Digital.HRAM.CapturedCycles
'''Theexec.Datalog.WriteComment ("Capture Cycle =" & CStr(numcap))
PinData = TheHdw.Digital.pins("SPI1_MISO").hram.PinPF(0, , numcap)

'added by rita
Dim site As Variant 'by rita
Dim sdbRegValue1 As New SiteDouble 'by rita
Dim sdbRegValue2 As New SiteDouble 'by rita
Dim sdbRegValue3 As New SiteDouble 'by rita
Dim pldMeasureValue As New PinListData
pldMeasureValue.AddPin "SPI1_MISO"
'end of addedj by rita

For Each site In TheExec.sites.Active
regdata = PinData.pins(0).value

For i = 0 To numcap - 1
If i Mod 4 = 0 Then
get_status_reg = get_status_reg + " "
End If

If regdata(i) = 2 Then
 get_status_reg = get_status_reg + "1"
Else
 get_status_reg = get_status_reg + "0"
End If
Next i

'comment by rita
'TheExec.Datalog.WriteComment ("Reg " + CStr(Hex(data)) + " , Binary Format: " + modify_data + ", Reg Value = " + get_status_reg)
'TheExec.Datalog.WriteComment ("")

'added by rita
Dim status_reg_ManufacturerID As Integer
Dim status_reg_MemoryType As Integer
Dim status_reg_MemoryCapacity As Integer
'20H
status_reg_ManufacturerID = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
sdbRegValue1.value = status_reg_ManufacturerID
'BBH
status_reg_MemoryType = (val(regdata(8)) - 1) * 2 ^ 7 + (val(regdata(9)) - 1) * 2 ^ 6 + (val(regdata(10)) - 1) * 2 ^ 5 + (val(regdata(11)) - 1) * 2 ^ 4 + (val(regdata(12)) - 1) * 2 ^ 3 + (val(regdata(13)) - 1) * 2 ^ 2 + (val(regdata(14)) - 1) * 2 ^ 1 + (val(regdata(15)) - 1) * 2 ^ 0
sdbRegValue2.value = status_reg_MemoryType
'19H
status_reg_MemoryCapacity = (val(regdata(16)) - 1) * 2 ^ 7 + (val(regdata(17)) - 1) * 2 ^ 6 + (val(regdata(18)) - 1) * 2 ^ 5 + (val(regdata(19)) - 1) * 2 ^ 4 + (val(regdata(20)) - 1) * 2 ^ 3 + (val(regdata(21)) - 1) * 2 ^ 2 + (val(regdata(22)) - 1) * 2 ^ 1 + (val(regdata(23)) - 1) * 2 ^ 0
sdbRegValue3.value = status_reg_MemoryCapacity
'end of added by rita

Next site
'added by rita
pldMeasureValue.pins("SPI1_MISO") = sdbRegValue1
'Set pldMeasureValue = sdbRegValue
'judgeCheckResult pldMeasureValue, 7000, "ManufacturerID"
 TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=32, hiVal:=32, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:="ManufacturerID", ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
' glb_TestInstance = theexec.DataManager.instancename
 'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
 pldMeasureValue.pins("SPI1_MISO") = sdbRegValue2
'judgeCheckResult pldMeasureValue, 7001, "MemoryType"
 TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=187, hiVal:=187, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:="MemoryType", ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
' glb_TestInstance = theexec.DataManager.instancename
' theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
 pldMeasureValue.pins("SPI1_MISO") = sdbRegValue3
'judgeCheckResult pldMeasureValue, 7002, "MemoryCapacity"
 TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=25, hiVal:=25, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:="MemoryCapacity", ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
' glb_TestInstance = theexec.DataManager.instancename
' theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    TheHdw.Utility.pins("K02").State = tlUtilBitOff
    TheHdw.Utility.pins("K04").State = tlUtilBitOff

'end of added by rita

End Function


Public Sub get_data(data As Integer, ByRef pat_modify() As String, data_len As Integer)
Dim i As Integer
For i = 0 To data_len - 1
If (data And 2 ^ (data_len - 1 - i)) <> 0 Then
pat_modify(i) = "1"
Else
pat_modify(i) = "0"
End If
Next i
End Sub



Public Function SetupHRam()
TheHdw.Digital.Patgen.Events.Clear
TheHdw.Digital.Patgen.Events.SetCycleCount True, 0, tlCycleTypeAbsolute
TheHdw.Digital.hram.CaptureType = captSTV
TheHdw.Digital.hram.SetTrigger trigFirst, False, 0
TheHdw.Digital.Patgen.HaltMode = 1
TheHdw.Digital.hram.size = 512
'theexec.Datalog.WriteComment ("Execute Hram Setup")
End Function

Public Function ReadEEPROM_Block()
    
       ' Dim iEEPROM As ExecDIBEEPROM
       ' Set iEEPROM = thehdw.DIB.EEPROM
        Dim REC() As IDIB_EEPROM_BlockObj
        ' EEPROMREC = iEEPROM.Block.List
         '  EEPROMREC = thehdw.DIB.EEPROM.Block.List
        REC = TheHdw.DIB.EEPROM.Block.list

        Dim tmpBoardNumber As Double
        tmpBoardNumber = REC(0).ID
        'tmpBoardNumber = thehdw.DIB.EEPROM.Block.List.ID
        Dim tmpBoardSize As Double
        tmpBoardSize = REC(0).size
       ' tmpBoardSize = thehdw.DIB.EEPROM.Block.List.size
        Dim vsite As Variant
        Dim pldMeasureValue As New PinListData
        Dim ResultsValue As Double
        Dim Tname As String
        Tname = "EEPROMRead"
           
        If tmpBoardNumber >= 0 And tmpBoardSize >= 0 Then
            ResultsValue = 1
        Else
            ResultsValue = 0
        End If
        
        pldMeasureValue.AddPin " "  'edit by rita on 2018May23'
        For Each vsite In TheExec.sites.Active '0521
            pldMeasureValue.pins(" ").value(vsite) = ResultsValue 'edit by rita on 2018May23'
        Next vsite '0521
       ' judgeCheckResult pldMeasureValue, 8004, "EEPROMRead", 1, 1
 TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=1, hiVal:=1, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
' glb_TestInstance = theexec.DataManager.instancename
' theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")

End Function

Public Function MCC_Cap_meas(pat As String, RX_Pinname As String, TX_Pinname As String, Lo_Limit As Double, Hi_Limit As Double, dblTestMeasureCurrent As Double, Validating_ As Boolean)
    Dim strTestPattern As String
    Dim strPatternPath As String
    strTestPattern = pat ' "CX_HSD_PCIE_UP_TX0_N.pat"
    strPatternPath = ".\Pattern\DIB\" ' & "Patterns\"
    
    If Validating_ Then
        Call PrLoadPattern(strPatternPath & strTestPattern)
        Exit Function    ' Exit after validation
    End If
    
    '***connect PPMU instrument***'
    'thehdw.PPMU.Pins("PCIE_UP_RX0_N,PCIE_UP_TX0_N").Connect
    TheHdw.Digital.pins(RX_Pinname).Disconnect
        TheHdw.Digital.pins(TX_Pinname).Disconnect
    TheHdw.PPMU.pins(RX_Pinname).Connect
    TheHdw.PPMU.pins(TX_Pinname).Connect
    '***-----------------------***'


    '***set PPMU instrument***'
    TheHdw.PPMU.pins(RX_Pinname).ForceV 0
    TheHdw.PPMU.pins(RX_Pinname).Gate = tlOn
    TheHdw.Wait 0.004
    '***-----------------------***'
    Dim Tname As String
    Tname = Replace(TheExec.DataManager.instancename, "_cap_value", "")
    '********************capacitor test begin********************'
    Dim strHSD_Pins As String
    Dim vsite As Variant
    Dim dblPeriod As Double
    Dim pldMeasureValue As New PinListData
    'Dim dblTestMeasureCurrent As Double
    Dim dblTestChargeVoltage As Double
    Dim dblTestOffsetVoltage As Double
    Dim pldFailCount As New PinListData
    Dim strPatternStartLabel As String


    'strHSD_Pins = "PCIE_UP_TX0_N"
    strHSD_Pins = TX_Pinname
    'dblTestMeasureCurrent = 0.0005
    'dblTestMeasureCurrent = 0.02
    dblTestChargeVoltage = 0.5
    dblTestOffsetVoltage = 0.3

    dblPeriod = 20 * ns
    strPatternStartLabel = "CX_003msPer020ns"

    strTestPattern = pat ' "CX_HSD_PCIE_UP_TX0_N.pat"
    strPatternPath = ".\Pattern\DIB\" ' & "Patterns\"

    With TheHdw.PPMU.pins(strHSD_Pins)
        .ForceV 0
        .Gate = tlOn
        TheHdw.Wait 0.005
        .Gate = tlOff
        TheHdw.Wait 0.002
        .Disconnect
    End With

    TheHdw.Digital.ApplyLevelsTiming False, True, True, , , , strHSD_Pins

    TheHdw.Digital.ConnectPins (strHSD_Pins)
    TheHdw.Wait 0.002
    With TheHdw.PinLevels.pins(strHSD_Pins)
        .ModifyLevel chVih, 2
        .ModifyLevel chVil, -0.5
        .ModifyLevel chVol, dblTestOffsetVoltage
        .ModifyLevel chVch, 5.5
        .ModifyLevel chVcl, -1
    End With

    TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVoh, dblTestOffsetVoltage + dblTestChargeVoltage

    TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIoh, 0
    TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIol, dblTestMeasureCurrent
    TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVt, 5.5

    TheHdw.Digital.ConnectPins (strHSD_Pins)
'    If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
'        TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Load
'    End If
'    TheHdw.Wait 0.002
'
'    If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
'        TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Load
'    End If
'     TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Load


    Dim Label_str As String

    Label_str = TX_Pinname & "_"
    
    If Label_str Like "GP_*" Then ' For Sicily 20200326
        Label_str = mid(Label_str, 4)
    ElseIf Label_str Like "ST_*" Then   ' For Sicily 20200526
        Label_str = Label_str
    End If
    
   ' thehdw.Digital.Patterns.Pat(strPatternPath & strTestPattern).Run "PCIE_UP_TX0_N_" & strPatternStartLabel
     TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Run Label_str & strPatternStartLabel
    TheHdw.Wait 0.005

    pldMeasureValue.AddPin (strHSD_Pins)
    pldFailCount = TheHdw.Digital.pins(strHSD_Pins).FailCount
    TheHdw.Wait 0.005

    For Each vsite In TheExec.sites.Active
        pldMeasureValue.pins(strHSD_Pins).value(vsite) = dblTestMeasureCurrent * ((pldFailCount.pins(strHSD_Pins).value(vsite) * dblPeriod) / dblTestChargeVoltage)
    Next vsite

    With TheHdw.PPMU.pins(strHSD_Pins)
        .ForceV 0
        TheHdw.Wait 0.02
        .Gate = tlOff
        TheHdw.Wait 0.002
    End With

    TheHdw.Digital.DisconnectPins (strHSD_Pins)

    '********************main test begin********************'
    'judgeCheckResult pldMeasureValue, 4808, "MCC9"
        TheExec.flow.TestLimit resultVal:=pldMeasureValue, PinName:=strHSD_Pins, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    If TheExec.sites.ActiveCount = 0 Then Exit Function
    '********************main test end********************'

'    If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
'        TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Unload
'    End If
'    TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Unload
    Set pldMeasureValue = Nothing
    '********************capacitor test end********************'

    '***disconnect PPMU instrument***'
    TheHdw.PPMU.pins(RX_Pinname).Gate = tlOff
    TheHdw.PPMU.pins(TX_Pinname).Gate = tlOff
    TheHdw.Wait 0.002
    TheHdw.PPMU.pins(RX_Pinname).Reset tlResetSettings
        TheHdw.PPMU.pins(TX_Pinname).Reset tlResetSettings
    TheHdw.PPMU.pins(RX_Pinname).Disconnect
      TheHdw.PPMU.pins(TX_Pinname).Disconnect
    '***-----------------------***'

End Function

Public Function LoopBack_1MHz_VOH_check(Validating_ As Boolean)

    Dim strTestPattern As String
    Dim strPatternPath As String
    Dim patTestPattern As String
    strTestPattern = "MeasFreq.pat"
    strPatternPath = ".\Pattern\DIB\" & strTestPattern
    patTestPattern = strPatternPath
'    patTestPattern.Value = strPatternPath
    Dim strTestPinlist As String
    Dim plstTestPinlist As New PinList
    
    Dim strTestPattern_waveform As String
    Dim strPatternPath_waveform As String
    Dim patTestPattern_waveform As New Pattern
    strTestPattern_waveform = "MeasFreq_waveform.pat"
    strPatternPath_waveform = ".\Pattern\DIB\" & strTestPattern_waveform
    patTestPattern_waveform.value = strPatternPath_waveform
    
    If Validating_ Then
        Call PrLoadPattern(patTestPattern)
        Exit Function    ' Exit after validation
    End If

    strTestPinlist = "PCIE_TX"
    TheHdw.Digital.pins("all_digital").Disconnect
    TheHdw.Digital.pins("PCIE_TX_RX").Connect
    'TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, PinLevelsSheet:="DIBC_Levels__", TimeSetSheet:="DIBC_TSets__"  '0831 by rita 'DIBC_TSets__
    '''add by rita on 2018Nov01, apply the pinlevel
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered  'add by rita on 2018Nov09 to connect digital pins
    TheHdw.Wait 0.01
'    thehdw.Digital.pins("PCIE_TX_RX").Levels.DriverMode = tlDriverModeLargeHiZ
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVil) = 0
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVih) = 1.2     '2
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVol) = 0.1
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVoh) = 0.8     '1.8
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chIol) = 0.001
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chIoh) = -0.001
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVt) = 1
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVcl) = -1
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVch) = 3
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.DriverMode = tlDriverModeLargeVt 'Original Largeswing-HiZ, Replace by Vt. It can eliminate overshoot in Pattern
    TheHdw.Wait 0.01
    
    TheHdw.Digital.Timing.period("DIBC_MeasFreq").value = 1 / (1 * 1000000)
    TheHdw.Digital.pins("PCIE_RX").Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (1 * 1000000) / 2
'    thehdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD1) = 1 / (1 * 1000000) / 2
'    thehdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (1 * 1000000)
    
    plstTestPinlist.value = strTestPinlist
    Call FreqSweptTest_1MHz(plstTestPinlist, 0.01, patTestPattern, 1000000, -0.2, 1.8, 0.01, 20)
'    Call FreqSweptTest_1MHz(plstTestPinlist, 0.002, patTestPattern, 1000000, 0.4, 0.6, 0.01, 20)
'    Call FreqSweptTest_waveform(plstTestPinlist, 0.002, patTestPattern_waveform, 1000000, -0.2, 1.8, 0.01, 20)
    TheHdw.Digital.pins("PCIE_TX_RX").Disconnect
'
    '''reset
    TheHdw.Wait 0.01
    TheHdw.Digital.ApplyLevelsTiming False, False, False
    TheHdw.Wait 0.01
    TheHdw.Digital.pins("PCIE_TX_RX").Disconnect
    If TheExec.sites.ActiveCount = 0 Then Exit Function
        
End Function

Public Function LoopBack_1MHz_VOH_check_2(Validating_ As Boolean)
    Dim strTestPattern As String
    Dim strPatternPath As String
    Dim patTestPattern As String
    strTestPattern = "MeasFreq_LPDPRX.pat"
    strPatternPath = ".\Pattern\DIB\" & strTestPattern
    patTestPattern = strPatternPath
    Dim strTestPinlist As String
    Dim plstTestPinlist As New PinList
    
    Dim strTestPattern_waveform As String
    Dim strPatternPath_waveform As String
    Dim patTestPattern_waveform As New Pattern
    strTestPattern_waveform = "MeasFreq_waveform.pat"
    strPatternPath_waveform = ".\Pattern\DIB\" & strTestPattern_waveform
    patTestPattern_waveform.value = strPatternPath_waveform
    
    If Validating_ Then
        Call PrLoadPattern(patTestPattern)
        Exit Function    ' Exit after validation
    End If

    strTestPinlist = "LPDPRX_RX_D_ODD"  'TX
    TheHdw.Digital.pins("all_digital").Disconnect
    TheHdw.Digital.pins("LPDPRX_RX_D").Connect
    'TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, PinLevelsSheet:="DIBC_Levels__", TimeSetSheet:="DIBC_TSets__"  '0831 by rita 'DIBC_TSets__
    '''add by rita on 2018Nov01, apply the pinlevel
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered  'add by rita on 2018Nov09 to connect digital pins
    TheHdw.Wait 0.01
'    thehdw.Digital.pins("LPDP_TX_RX").Levels.DriverMode = tlDriverModeLargeHiZ
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVil) = 0
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVih) = 1.2   '2 0.9
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVol) = 0.1
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVoh) = 0.8    '1.8 0.8
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chIol) = 0.001
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chIoh) = -0.001
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVt) = 1
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVcl) = -1
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVch) = 3
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.DriverMode = tlDriverModeLargeVt  'Original Largeswing-HiZ, Replace by Vt. It can eliminate overshoot in Pattern
    TheHdw.Wait 0.01
    
    TheHdw.Digital.Timing.period("DIBC_MeasFreq_2").value = 1 / (1 * 1000000)
    TheHdw.Digital.pins("LPDPRX_RX_D_EVEN").Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD2) = 1 / (1 * 1000000) / 2  'RX
'    thehdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD1) = 1 / (1 * 1000000) / 2
'    thehdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (1 * 1000000)
    
    plstTestPinlist.value = strTestPinlist
    Call FreqSweptTest_1MHz(plstTestPinlist, 0.01, patTestPattern, 1000000, -0.2, 1.8, 0.01, 20)
'    Call FreqSweptTest_1MHz(plstTestPinlist, 0.002, patTestPattern, 1000000, 0.4, 0.6, 0.01, 20)
'    Call FreqSweptTest_waveform(plstTestPinlist, 0.002, patTestPattern_waveform, 1000000, -0.2, 1.8, 0.01, 20)
    TheHdw.Digital.pins("LPDPRX_RX_D").Disconnect
'
    '''reset
    TheHdw.Wait 0.01
    TheHdw.Digital.ApplyLevelsTiming False, False, False
    TheHdw.Wait 0.01
    TheHdw.Digital.pins("LPDPRX_RX_D").Disconnect
    If TheExec.sites.ActiveCount = 0 Then Exit Function
        
End Function

Public Function FreqSweptTest_1MHz(meas_pin As PinList, MeasInterval As Double, pat As String, MinFreq As Double, StartVolt As Double, StopVolt As Double, Resolution As Double, Optional MaxCheckTime As Integer = 20)
'   On Error GoTo errhandler
   
    Dim group_PCIE_RX As String
    Dim group_PCIE_TX As String
    group_PCIE_RX = "PCIE_RX0_P,PCIE_RX0_N,PCIE_RX1_P,PCIE_RX1_N,PCIE_RX2_P,PCIE_RX2_N,PCIE_RX3_P,PCIE_RX3_N,PCIE_RX4_P,PCIE_RX4_N"
    group_PCIE_TX = "PCIE_TX0_P,PCIE_TX0_N,PCIE_TX1_P,PCIE_TX1_N,PCIE_TX2_P,PCIE_TX2_N,PCIE_TX3_P,PCIE_TX3_N,PCIE_TX4_P,PCIE_TX4_N"
    
    Dim LowerVoh As New PinListData
    Dim UpperVoh As New PinListData
    Dim VohSwept As New PinListData
    Dim MidVolt As Double
    MidVolt = StartVolt + (StopVolt - StartVolt) / 2
    
    
    
'    TheHdw.Patterns(pat).Load

'''    thehdw.Patterns(pat).Start
'''
'''    If (TheExec.TesterMode = testModeOnline) Then
'''       thehdw.Digital.Patgen.FlagWait cpuA, 0
'''    End If
    
    MeasFreq_VohSwept meas_pin.value, MeasInterval, StartVolt, MidVolt, Resolution, MinFreq, LowerVoh, pat, MaxCheckTime
    MeasFreq_VohSwept meas_pin.value, MeasInterval, MidVolt, StopVolt, Resolution, MinFreq, UpperVoh, pat, MaxCheckTime
    VohSwept = UpperVoh.Math.Subtract(LowerVoh)
    
'''    Call thehdw.Digital.Patgen.Continue(0, cpuA) 'clean cpuA
    'TheExec.Flow.TestLimit resultVal:=VohSwept, unit:=unitVolt, Tname:="Clock Amp", ForceResults:=tlForceNone 'comment by rita on 2018Nov09
    '********************print datalog********************'
    Dim LowLimit As Double
    Dim hiLimit As Double
    Dim Tname As String
    LowLimit = 0.9
    hiLimit = 1.1
    Tname = "LoopBack_1Mhz_Voh"
    
    VohSwept = VohSwept.Math.Multiply(2)    'Replace HiZ by Vt, swing need to be multiply by 2
    
    Dim pldMeasureValue As New PinListData
    Set pldMeasureValue = VohSwept
      TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitAmp, customUnit:="V", ForceResults:=tlForceNone
      'glb_TestInstance = theexec.DataManager.instancename
      'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    'judgeCheckResult pldMeasureValue, 100, "Clock Amp_1MHz" ', 2, 3, "V"
    If TheExec.sites.ActiveCount = 0 Then Exit Function
    '****************************************************'
    
    If (TheExec.TesterMode = testModeOnline) Then
        TheHdw.Digital.Patgen.Continue 0, cpuA
         'thehdw.Wait 0.01
        If (TheHdw.Digital.Patgen.IsRunning = True) Then
         TheHdw.Digital.Patgen.Halt
        End If
    End If
    
   
    Exit Function
errHandler:
    MsgBox "Error encountered in FreqTest " + vbCrLf + _
        "VBT Error # " + Trim(str(err.number)) + ": " + err.Description
End Function


Public Function MeasFreq_VohSwept(meas_pin As String, MeasInterval As Double, StartPoint As Double, StopPoint As Double, Resolution As Double _
, MinFreq As Double, MeasData As PinListData, pat As String, Optional MaxCheckTime As Integer = 20)
    
'    On Error Resume Next
    Dim LowerPointData As ExecData
    Dim UpperPointData As ExecData
    Dim FirstPointData As ExecData
    ''Dim MeasData As New PinListData
    Dim NextVohPoint As New PinListData
    Dim site As Variant
    Dim stopwhile As Boolean

    Dim mExecResult() As ExecData
    Dim nExecData As ExecData
    Dim Pin_Arr1() As String
    Dim PinCount As Long
    Dim ExecNum As Integer
    Dim PinNum As Integer
    Dim i As Integer
    
    ReDim mExecResult(MaxCheckTime - 1)
    TheExec.DataManager.DecomposePinList meas_pin, Pin_Arr1(), PinCount
    InitPinListData Pin_Arr1, MeasData
    InitExecData Pin_Arr1, LowerPointData
    InitExecData Pin_Arr1, UpperPointData
    InitExecData Pin_Arr1, FirstPointData
    For i = 0 To UBound(mExecResult)
       InitExecData Pin_Arr1, mExecResult(i)
    Next i
    ''''' Init Data
    For PinNum = 0 To UBound(LowerPointData.CurrMeasData)
        For Each site In TheExec.sites.Active
'           LowerPointData.CurrMeasData(PinNum).Threshold(Site) = StartPoint + Site + PinNum * 0.1 'comment by rita on 2018NOv08
'           UpperPointData.CurrMeasData(PinNum).Threshold(Site) = StopPoint + Site + PinNum * 0.1
           LowerPointData.CurrMeasData(PinNum).Threshold(site) = StartPoint '+ PinNum * 0.1  'comment by rita on 2018Nov09
           UpperPointData.CurrMeasData(PinNum).Threshold(site) = StopPoint '+ PinNum * 0.1 'comment by rita on 2018Nov09
           FirstPointData.CurrMeasData(PinNum).Threshold(site) = (LowerPointData.CurrMeasData(PinNum).Threshold(site) + _
           UpperPointData.CurrMeasData(PinNum).Threshold(site)) / 2
        Next site
    Next PinNum

    Call MeasFreqFunc(meas_pin, MeasInterval, LowerPointData, MinFreq, pat, 100)
    Call MeasFreqFunc(meas_pin, MeasInterval, UpperPointData, MinFreq, pat, 0)
    Call MeasFreqFunc(meas_pin, MeasInterval, FirstPointData, MinFreq, pat, 50)
    
    mExecResult(0).PrevLowerData = LowerPointData.CurrMeasData
    mExecResult(0).PrevUpperData = UpperPointData.CurrMeasData
    mExecResult(0).CurrMeasData = FirstPointData.CurrMeasData
    If (PrintDebug = True) Then
      PrintDebugLog mExecResult(0), 0
    End If
    stopwhile = CheckStopCondition(mExecResult(0))
    ExecNum = 1
    While (stopwhile = False And ExecNum < MaxCheckTime)
       FindNextExecDef mExecResult(ExecNum), mExecResult(ExecNum - 1), Resolution
       Call MeasFreqFunc(meas_pin, MeasInterval, mExecResult(ExecNum), MinFreq, pat, 50)
       stopwhile = CheckStopCondition(mExecResult(ExecNum))
       If (PrintDebug = True) Then
         PrintDebugLog mExecResult(ExecNum), ExecNum
       End If
       
       If (stopwhile = False) Then
          ExecNum = ExecNum + 1
       End If
    Wend
    
    If TheExec.TesterMode = testModeOnline Then 'AL20181130, add only execute online, offline will cause errors.
        ReDim Preserve mExecResult(ExecNum)  '''bug found by rita on 2018Nov08
    End If
    Dim FinalExecData As ExecData
    FinalExecData = mExecResult(UBound(mExecResult)) ' if occur error, please extended StartPoint&StopPoint voltage

    For PinNum = 0 To UBound(FinalExecData.CurrMeasData)
      For Each site In TheExec.sites.Active
       If (FinalExecData.CurrMeasData(PinNum).FinishSearch(site) = True) Then
         MeasData.pins(FinalExecData.CurrMeasData(PinNum).PinName).value(site) = FinalExecData.CurrMeasData(PinNum).Threshold(site)
       Else
         MeasData.pins(FinalExecData.CurrMeasData(PinNum).PinName).value(site) = -1
       End If
      Next site
    Next PinNum
    
End Function

Public Function MeasFreqFunc(mea_pin As String, MeasInterval As Double, MeasResult As ExecData, MinFreq As Double, pat As String, Optional OfflineData As Integer = 100) As Long
    Dim freqCnt1 As New PinListData
    Dim PinFreq As New PinListData
    Dim site As Variant
    Dim PinNum As Integer
    'thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered   'comment by rita on 2018Nov01
    For PinNum = 0 To UBound(MeasResult.CurrMeasData)
        For Each site In TheExec.sites.Active
            TheHdw.Digital.pins(MeasResult.CurrMeasData(PinNum).PinName).Levels.value(chVoh) = MeasResult.CurrMeasData(PinNum).Threshold(site)
            TheHdw.Digital.pins(MeasResult.CurrMeasData(PinNum).PinName).Levels.value(chVol) = MeasResult.CurrMeasData(PinNum).Threshold(site)
            'Debug.Print MeasResult.CurrMeasData(pinnum).Threshold(Site) & "_VOH"
        Next site
    Next PinNum
    
    TheHdw.patterns(pat).start
    

    If (TheExec.TesterMode = testModeOnline) Then
       TheHdw.Digital.Patgen.FlagWait cpuA, 0
    End If
    
    
    With TheHdw.Digital.pins(mea_pin).FreqCtr
        .Clear
        .EventSource = VOH 'sweep the level
        .EventSlope = Positive
        .Interval = MeasInterval
        .Enable = IntervalEnable
    End With
    TheHdw.Digital.pins(mea_pin).FreqCtr.start
    freqCnt1 = TheHdw.Digital.pins(mea_pin).FreqCtr.Read
    ''''Offline Data
    
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)
    TheHdw.Digital.Patgen.HaltWait
    If (TheExec.TesterMode = testModeOffline) Then
       FillSimulateData freqCnt1, OfflineData
    End If
    ''''''
    PinFreq = freqCnt1.Math.divide(MeasInterval)
    'Debug.Print PinFreq.pins(0).Value(0) & "_MeasFreq"
    For PinNum = 0 To UBound(MeasResult.CurrMeasData)
       For Each site In TheExec.sites.Active
           MeasResult.CurrMeasData(PinNum).MeasVal = PinFreq.pins(MeasResult.CurrMeasData(PinNum).PinName).value(site)
           If (MeasResult.CurrMeasData(PinNum).MeasVal(site) >= MinFreq) Then
               MeasResult.CurrMeasData(PinNum).MeasPF(site) = 2 'pass, measured expected freq 1MHZ
           Else
               MeasResult.CurrMeasData(PinNum).MeasPF(site) = 1 'fail, not measured the expected freq 1MHZ
           End If
       Next site
    Next PinNum
End Function


Public Function CheckStopCondition(nExecData As ExecData) As Boolean
  Dim PinNum As Integer
  Dim site As Variant
  Dim StopCheck As Boolean
  CheckStopCondition = False
  StopCheck = False
  For PinNum = 0 To UBound(nExecData.CurrMeasData)
    For Each site In TheExec.sites.Active
        If (nExecData.CurrMeasData(PinNum).FinishSearch(site) = False) Then
          StopCheck = True
        End If
    Next site
    If (StopCheck = True) Then
      Exit For
    End If
  Next PinNum
CheckStopCondition = Not StopCheck
End Function

Public Function InitPinListData(pins() As String, PinsData As PinListData)
  ''Dim Pin_Arr1() As String
  ''Dim PinCount As Long
  Dim PinNum As Integer
  Dim site As Variant
  ''theexec.DataManager.DecomposePinList pins, Pin_Arr1(), PinCount
  For PinNum = 0 To UBound(pins)
     PinsData.AddPin (pins(PinNum))
     For Each site In TheExec.sites.Active
        PinsData.pins(PinNum).value(site) = -1
     Next site
  Next PinNum
End Function

Public Function InitExecData(pins() As String, nExecData As ExecData)
  Dim Pin_Arr1() As String
  Dim PinCount As Long
  Dim PinNum As Integer
  Dim site As Variant
  ''theexec.DataManager.DecomposePinList pins, Pin_Arr1(), PinCount
  
  ReDim nExecData.CurrMeasData(UBound(pins))
  ReDim nExecData.PrevLowerData(UBound(pins))
  ReDim nExecData.PrevUpperData(UBound(pins))
  
  For PinNum = 0 To UBound(pins)
        nExecData.CurrMeasData(PinNum).PinName = pins(PinNum)
        nExecData.PrevLowerData(PinNum).PinName = pins(PinNum)
        nExecData.PrevUpperData(PinNum).PinName = pins(PinNum)
     ''For Each Site In TheExec.Sites.Active
        nExecData.CurrMeasData(PinNum).FinishSearch = False
        nExecData.CurrMeasData(PinNum).MeasPF = 0
        nExecData.CurrMeasData(PinNum).MeasVal = 0
        nExecData.CurrMeasData(PinNum).Threshold = 0
        
        nExecData.PrevLowerData(PinNum).FinishSearch = False
        nExecData.PrevLowerData(PinNum).MeasPF = 0
        nExecData.PrevLowerData(PinNum).MeasVal = 0
        nExecData.PrevLowerData(PinNum).Threshold = 0
        
        nExecData.PrevUpperData(PinNum).FinishSearch = False
        nExecData.PrevUpperData(PinNum).MeasPF = 0
        nExecData.PrevUpperData(PinNum).MeasVal = 0
        nExecData.PrevUpperData(PinNum).Threshold = 0
        
     ''Next Site
  Next PinNum
End Function

Public Function PrintDebugLog(MeasData As ExecData, ExecCount As Integer)
  Dim PinNum As Integer
  Dim site As Variant
  Dim PrintStr As String
  TheExec.Datalog.WriteComment (" ========== " + CStr(ExecCount) + " ========== ")
    PrintStr = ExtendStr("PinName") + ExtendStr("Type") + ExtendStr("Site", 5) + _
    ExtendStr("Threshold") + ExtendStr("MeasVal") + _
    ExtendStr("P/F") + ExtendStr("Finish")
    TheExec.Datalog.WriteComment (PrintStr)
  
  
  For PinNum = 0 To UBound(MeasData.CurrMeasData)
    For Each site In TheExec.sites.Active
       PrintStr = ExtendStr(MeasData.CurrMeasData(PinNum).PinName) + ExtendStr("Current") + ExtendStr(CStr(site), 5) + _
       ExtendStr(CStr(MeasData.CurrMeasData(PinNum).Threshold(site))) + ExtendStr(CStr(MeasData.CurrMeasData(PinNum).MeasVal(site))) + _
       ExtendStr(CStr(MeasData.CurrMeasData(PinNum).MeasPF(site))) + ExtendStr(CStr(MeasData.CurrMeasData(PinNum).FinishSearch(site)))
       TheExec.Datalog.WriteComment (PrintStr)
       
       PrintStr = ExtendStr(MeasData.PrevLowerData(PinNum).PinName) + ExtendStr("PrevLowerData") + ExtendStr(CStr(site), 5) + _
       ExtendStr(CStr(MeasData.PrevLowerData(PinNum).Threshold(site))) + ExtendStr(CStr(MeasData.PrevLowerData(PinNum).MeasVal(site))) + _
       ExtendStr(CStr(MeasData.PrevLowerData(PinNum).MeasPF(site))) + ExtendStr(CStr(MeasData.PrevLowerData(PinNum).FinishSearch(site)))
       TheExec.Datalog.WriteComment (PrintStr)
       
        PrintStr = ExtendStr(MeasData.PrevUpperData(PinNum).PinName) + ExtendStr("PrevUpperData") + ExtendStr(CStr(site), 5) + _
       ExtendStr(CStr(MeasData.PrevUpperData(PinNum).Threshold(site))) + ExtendStr(CStr(MeasData.PrevUpperData(PinNum).MeasVal(site))) + _
       ExtendStr(CStr(MeasData.PrevUpperData(PinNum).MeasPF(site))) + ExtendStr(CStr(MeasData.PrevUpperData(PinNum).FinishSearch(site)))
       TheExec.Datalog.WriteComment (PrintStr)
    Next site
  Next PinNum
End Function

Function ExtendStr(data As String, Optional exp_length As Integer = 20) As String
   Dim strLen As Integer
   Dim i As Integer
   Dim NewStr As String
   strLen = Len(data)
   If (strLen < exp_length) Then
     For i = 0 To exp_length - strLen
       NewStr = NewStr + " "
     Next i
     NewStr = NewStr + data
   Else
     NewStr = data
   End If
   ExtendStr = NewStr
End Function
Public Function FindNextExecDef(nExecData As ExecData, PrevExec As ExecData, Resolution As Double)
    Dim PinNum As Integer
    Dim site As Variant
    For PinNum = 0 To UBound(PrevExec.CurrMeasData)
      For Each site In TheExec.sites.Active
        If (PrevExec.CurrMeasData(PinNum).FinishSearch(site) = True) Then
           AssignPinMeasData nExecData.PrevLowerData(PinNum), PrevExec.PrevLowerData(PinNum), site
           AssignPinMeasData nExecData.PrevUpperData(PinNum), PrevExec.PrevUpperData(PinNum), site
           AssignPinMeasData nExecData.CurrMeasData(PinNum), PrevExec.CurrMeasData(PinNum), site
        Else
            If (PrevExec.CurrMeasData(PinNum).MeasPF(site) <> PrevExec.PrevLowerData(PinNum).MeasPF(site)) Then
                AssignPinMeasData nExecData.PrevUpperData(PinNum), PrevExec.CurrMeasData(PinNum), site
                AssignPinMeasData nExecData.PrevLowerData(PinNum), PrevExec.PrevLowerData(PinNum), site
            ElseIf (PrevExec.CurrMeasData(PinNum).MeasPF(site) <> PrevExec.PrevUpperData(PinNum).MeasPF(site)) Then
                AssignPinMeasData nExecData.PrevUpperData(PinNum), PrevExec.PrevUpperData(PinNum), site
                AssignPinMeasData nExecData.PrevLowerData(PinNum), PrevExec.CurrMeasData(PinNum), site
            Else
                AssignPinMeasData nExecData.PrevUpperData(PinNum), PrevExec.PrevUpperData(PinNum), site
                AssignPinMeasData nExecData.PrevLowerData(PinNum), PrevExec.PrevLowerData(PinNum), site
                AssignPinMeasData nExecData.CurrMeasData(PinNum), PrevExec.CurrMeasData(PinNum), site
            End If
    
            If (Math.Abs((nExecData.PrevLowerData(PinNum).Threshold(site) - nExecData.PrevUpperData(PinNum).Threshold(site))) > Resolution) Then
                nExecData.CurrMeasData(PinNum).Threshold(site) = (nExecData.PrevLowerData(PinNum).Threshold(site) + _
                nExecData.PrevUpperData(PinNum).Threshold(site)) / 2
            Else
                AssignPinMeasData nExecData.CurrMeasData(PinNum), PrevExec.CurrMeasData(PinNum), site
                nExecData.CurrMeasData(PinNum).FinishSearch(site) = True

            End If
        End If
      Next site
    Next PinNum
End Function

Public Function FillSimulateData(MeasData As PinListData, OfflineData As Integer)
    Dim PinNum As Integer
    Dim rndvalue As Integer
    Dim site As Variant
        
    For PinNum = 0 To MeasData.pins.Count - 1
     For Each site In TheExec.sites.Active
        rndvalue = Int(((OfflineData + 10) - OfflineData) * Rnd + OfflineData)
        MeasData.pins(PinNum).value(site) = rndvalue
     Next site
    Next PinNum

End Function

Public Function AssignPinMeasData(TargetData As PinMeasData, SrcData As PinMeasData, site As Variant)
    TargetData.Threshold(site) = SrcData.Threshold(site)
    TargetData.FinishSearch(site) = SrcData.FinishSearch(site)
    TargetData.MeasPF(site) = SrcData.MeasPF(site)
    TargetData.MeasVal(site) = SrcData.MeasVal(site)
End Function

Public Function LoopBack_Freq(Validating_ As Boolean)
    On Error GoTo errHandler
    
    Dim CountVal As New PinListData
    Dim VihVal As Double
    Dim VodVal As Double
    Dim site As Variant
    Dim pathPathStr_freq As String
    Dim rx_p As String
    Dim rx_n As String
    Dim tx_diff As String
    Dim rx_pins As String
    Dim tx_pins As String
    Dim Lo_Limit As Double
    Dim Hi_Limit As Double
    Dim Tname As String
    
    rx_p = "GP_PCIE_RX0_P,GP_PCIE_RX1_P,GP_PCIE_RX2_P,ST_PCIE_RX0_P"
    rx_n = "GP_PCIE_RX0_N,GP_PCIE_RX1_N,GP_PCIE_RX2_N,ST_PCIE_RX0_N"
    tx_diff = "PCIE_TX0_DIFF,PCIE_TX1_DIFF,PCIE_TX2_DIFF,ST_PCIE_TX0_DIFF"
    rx_pins = "GP_PCIE_RX0_P,GP_PCIE_RX0_N,GP_PCIE_RX1_P,GP_PCIE_RX1_N,GP_PCIE_RX2_P,GP_PCIE_RX2_N,ST_PCIE_RX0_P,ST_PCIE_RX0_N"

    pathPathStr_freq = ".\Pattern\DIB\DIBC_loopFreqMeas.pat"
    
    If Validating_ Then
        Call PrLoadPattern(pathPathStr_freq)
        Exit Function    ' Exit after validation
    End If
    
'    thehdw.Digital.pins(rx_pins & "," & tx_diff).Connect
    TheHdw.Digital.pins("all_digital").Disconnect
    TheHdw.Digital.pins("PCIE_TX_RX").Connect
    TheHdw.Digital.ApplyLevelsTiming True, True, True
    
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVil) = 0
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVih) = 3
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVol) = 0
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVoh) = 0.4
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chIol) = 0
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chIoh) = 0
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVt) = 0.4
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVcl) = -1
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVch) = 6
'''    thehdw.Digital.pins(rx_pins).Levels.DriverMode = tlDriverModeVt
'''
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVicm) = 0.6
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVid) = 0.6
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVid0) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVid1) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVicm0) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVicm1) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVod) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVod_alt1) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVod_alt2) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVod0) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVod1) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Iol) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Ioh) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiffVodTyp) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiffVocmTyp) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Vt) = 0.6
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Vcl) = -1
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Vch) = 6
'''    thehdw.Digital.pins(tx_diff).Levels.DriverMode = tlDriverModeVt
'''    thehdw.Wait 0.01
    
'    thehdw.Digital.pins("PCIE_TX_RX").Levels.DriverMode = tlDriverModeLargeHiZ
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVil) = 0
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVih) = 1.5 'Sicily update from 1v to 1.5v
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVol) = 0.4
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVoh) = 0.4
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chIol) = 0
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chIoh) = 0
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVt) = 0.4
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVcl) = -1
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.value(chVch) = 6
    TheHdw.Digital.pins("PCIE_TX_RX").Levels.DriverMode = tlDriverModeVt
    TheHdw.Wait 0.01
    
'
    
'    TheHdw.Patterns(pathPathStr_freq).Load

    '***********************62.5MHZ start*****************************
    TheHdw.Digital.Timing.period("DIBC_MeasFreq").value = 1 / (62.5 * 1000000)

    TheHdw.Digital.pins(rx_p).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (62.5 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD1) = 1 / (62.5 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (62.5 * 1000000)
    
    TheHdw.Digital.Timing.ApplyTimingChanges True

    TheHdw.patterns(pathPathStr_freq).start

    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    
    TheHdw.Wait 0.1
    
    With TheHdw.Digital.pins(tx_diff).FreqCtr 'tx_diff
        .Clear
        .EventSlope = Positive
        .EventSource = vol
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With
    
    CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    CountVal = CountVal.Math.divide(TheHdw.Digital.pins(tx_diff).FreqCtr.Interval)
        
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)

    TheHdw.Digital.Patgen.HaltWait
    
    VihVal = TheHdw.Digital.pins("GP_PCIE_RX0_P").Levels.value(chVih)
    
'''    Theexec.Datalog.WriteComment "Programmed Vih : " + CStr(VihVal)
    
    Dim li As Long
    Dim vsite As Variant
    Dim pldPosMeas As New PinListData
    For li = 0 To 7
        If right(CountVal.pins(li).name, 2) = "_n" Then

        ElseIf right(CountVal.pins(li).name, 2) = "_p" Then
            pldPosMeas.AddPin CountVal.pins(li).name
            For Each vsite In TheExec.sites.Active
                pldPosMeas = CountVal.pins(li).value(vsite)
            Next vsite
        End If
                
    Next li
 '  judgeCheckResult pldPosMeas, 300, "LoopBack_62.5MHZ"
 Lo_Limit = 61500000
 Hi_Limit = 63500000
 Tname = "LoopBack_62.5MHz"
      TheExec.flow.TestLimit resultVal:=pldPosMeas, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="", ForceResults:=tlForceNone
      'glb_TestInstance = theexec.DataManager.instancename
      'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
   If TheExec.sites.ActiveCount = 0 Then Exit Function
'    theexec.Flow.TestLimit pldPosMeas, 50000000#, 75000000#, , , , , , "loopback_62.5MHZ", , , , , , , , 10000
    '***********************62.5MHZ end*******************************
    
    '***********************125MHZ start******************************
    TheHdw.Digital.Timing.period("DIBC_MeasFreq").value = 1 / (125 * 1000000)
    TheHdw.Digital.pins(rx_p).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (125 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD1) = 1 / (125 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (125 * 1000000)

    TheHdw.Digital.Timing.ApplyTimingChanges True

    TheHdw.patterns(pathPathStr_freq).start

    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    
    TheHdw.Wait 0.1
    
    With TheHdw.Digital.pins(tx_diff).FreqCtr
        .Clear
        .EventSlope = Positive
        .EventSource = vol
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With
    
    CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    CountVal = CountVal.Math.divide(TheHdw.Digital.pins("PCIE_TX").FreqCtr.Interval)
        
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)

    TheHdw.Digital.Patgen.HaltWait
    
    VihVal = TheHdw.Digital.pins("GP_PCIE_RX0_P").Levels.value(chVih)
    
'''    Theexec.Datalog.WriteComment "Programmed Vih : " + CStr(VihVal)
    
    Set pldPosMeas = Nothing 'clean data
    For li = 0 To 7
        If right(CountVal.pins(li).name, 2) = "_n" Then

        ElseIf right(CountVal.pins(li).name, 2) = "_p" Then
            pldPosMeas.AddPin CountVal.pins(li).name
            For Each vsite In TheExec.sites.Active
                pldPosMeas = CountVal.pins(li).value(vsite)
            Next vsite
        End If
                
    Next li
  ' judgeCheckResult pldPosMeas, 304, "LoopBack_125MHZ"
    Lo_Limit = 124000000
 Hi_Limit = 126000000
 Tname = "LoopBack_125MHz"
      TheExec.flow.TestLimit resultVal:=pldPosMeas, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="", ForceResults:=tlForceNone
      'glb_TestInstance = theexec.DataManager.instancename
      'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
   If TheExec.sites.ActiveCount = 0 Then Exit Function
'    theexec.Flow.TestLimit pldPosMeas, 100000000#, 150000000#, , , , , , "loopback_125MHZ"
    '***********************125MHZ end********************************
    
    '***********************250MHZ start******************************
    TheHdw.Digital.Timing.period("DIBC_MeasFreq").value = 1 / (200 * 1000000)
    TheHdw.Digital.pins(rx_p).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (200 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD1) = 1 / (200 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq", chEdgeD2) = 1 / (200 * 1000000)
    
    TheHdw.Digital.Timing.ApplyTimingChanges True

    TheHdw.patterns(pathPathStr_freq).start

    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    
    TheHdw.Wait 0.2
    
    With TheHdw.Digital.pins(tx_diff).FreqCtr
        .Clear
        .EventSlope = Positive
        .EventSource = vol
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With
    
    CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    CountVal = CountVal.Math.divide(TheHdw.Digital.pins("PCIE_TX").FreqCtr.Interval)
        
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)

    TheHdw.Digital.Patgen.HaltWait
    
    VihVal = TheHdw.Digital.pins("GP_PCIE_RX0_P").Levels.value(chVih)
    
'''    Theexec.Datalog.WriteComment "Programmed Vih : " + CStr(VihVal)
    
    Set pldPosMeas = Nothing 'clean data
    For li = 0 To 7
        If right(CountVal.pins(li).name, 2) = "_n" Then

        ElseIf right(CountVal.pins(li).name, 2) = "_p" Then
            pldPosMeas.AddPin CountVal.pins(li).name
            For Each vsite In TheExec.sites.Active
                pldPosMeas = CountVal.pins(li).value(vsite)
            Next vsite
        End If
                
    Next li
   'judgeCheckResult pldPosMeas, 308, "LoopBack_200MHZ"
       Lo_Limit = 198000000
 Hi_Limit = 202000000
 Tname = "LoopBack_200MHz"
      TheExec.flow.TestLimit resultVal:=pldPosMeas, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="", ForceResults:=tlForceNone
      'glb_TestInstance = theexec.DataManager.instancename
      'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
   If TheExec.sites.ActiveCount = 0 Then Exit Function
'    theexec.Flow.TestLimit pldPosMeas, 160000000#, 240000000#, , , , , , "loopback_250MHZ"
    '***********************250MHZ end********************************
    TheHdw.Digital.pins("PCIE_TX_RX").Disconnect
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function LoopBack_Freq_2(Validating_ As Boolean) ' Validating_ As Boolean
    On Error GoTo errHandler
    
    Dim CountVal As New PinListData
    Dim VihVal As Double
    Dim VodVal As Double
    Dim site As Variant
    Dim pathPathStr_freq As String
    Dim rx_p As String
    Dim rx_n As String
    Dim tx_diff As String
    Dim rx_pins As String
    Dim tx_pins As String
    Dim Lo_Limit As Double
    Dim Hi_Limit As Double
    Dim Tname As String
    
    rx_p = "LPDPRX_RX_D0_P,LPDPRX_RX_D2_P,LPDPRX_RX_D4_P,LPDPRX_RX_D6_P,LPDPRX_RX_D8_P,LPDPRX_RX_D10_P" '' "GP_PCIE_RX0_P,GP_PCIE_RX1_P,GP_PCIE_RX2_P,ST_PCIE_RX0_P"
    rx_n = "LPDPRX_RX_D0_N,LPDPRX_RX_D2_N,LPDPRX_RX_D4_N,LPDPRX_RX_D6_N,LPDPRX_RX_D8_N,LPDPRX_RX_D10_N" '' "GP_PCIE_RX0_N,GP_PCIE_RX1_N,GP_PCIE_RX2_N,ST_PCIE_RX0_N"
    tx_diff = "LPDPRX_RX_D1_DIFF,LPDPRX_RX_D3_DIFF,LPDPRX_RX_D5_DIFF,LPDPRX_RX_D7_DIFF,LPDPRX_RX_D9_DIFF,LPDPRX_RX_D11_DIFF" '' "PCIE_TX0_DIFF,PCIE_TX1_DIFF,PCIE_TX2_DIFF,ST_PCIE_TX0_DIFF"
    rx_pins = rx_p & rx_n '' "GP_PCIE_RX0_P,GP_PCIE_RX0_N,GP_PCIE_RX1_P,GP_PCIE_RX1_N,GP_PCIE_RX2_P,GP_PCIE_RX2_N,ST_PCIE_RX0_P,ST_PCIE_RX0_N"

    pathPathStr_freq = ".\Pattern\DIB\DIBC_loopFreqMeas_LPDPRX.pat"
    
    If Validating_ Then
        Call PrLoadPattern(pathPathStr_freq)
        Exit Function    ' Exit after validation
    End If
    
'    thehdw.Digital.pins(rx_pins & "," & tx_diff).Connect
    TheHdw.Digital.pins("all_digital").Disconnect
    TheHdw.Digital.pins("LPDPRX_RX_D").Connect
    TheHdw.Digital.ApplyLevelsTiming True, True, True
    
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVil) = 0
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVih) = 3
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVol) = 0
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVoh) = 0.4
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chIol) = 0
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chIoh) = 0
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVt) = 0.4
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVcl) = -1
'''    thehdw.Digital.pins(rx_pins).Levels.Value(chVch) = 6
'''    thehdw.Digital.pins(rx_pins).Levels.DriverMode = tlDriverModeVt
'''
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVicm) = 0.6
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVid) = 0.6
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVid0) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVid1) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVicm0) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVicm1) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVod) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVod_alt1) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chVod_alt2) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVod0) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chdVod1) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Iol) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Ioh) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiffVodTyp) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiffVocmTyp) = 0
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Vt) = 0.6
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Vcl) = -1
'''    thehdw.Digital.pins(tx_diff).DifferentialLevels.Value(chDiff_Vch) = 6
'''    thehdw.Digital.pins(tx_diff).Levels.DriverMode = tlDriverModeVt
'''    thehdw.Wait 0.01
    
'    thehdw.Digital.pins("PCIE_TX_RX").Levels.DriverMode = tlDriverModeLargeHiZ
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVil) = 0
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVih) = 2#  'Sicily update from 1v to 1.5v
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVol) = 0.4
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVoh) = 0.4
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chIol) = 0
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chIoh) = 0
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVt) = 0.4
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVcl) = -1
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.value(chVch) = 6
    TheHdw.Digital.pins("LPDPRX_RX_D").Levels.DriverMode = tlDriverModeVt
    TheHdw.Wait 0.01
    
'
'''''    If Validating_ Then
'''''        Call PrLoadPattern(pathPathStr_freq)
'''''        Exit Function    ' Exit after validation
'''''    End If
    
'    TheHdw.Patterns(pathPathStr_freq).Load

    '***********************62.5MHZ start*****************************
    TheHdw.Digital.Timing.period("DIBC_MeasFreq_2").value = 1 / (62.5 * 1000000)

    TheHdw.Digital.pins(rx_p).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD2) = 1 / (62.5 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD1) = 1 / (62.5 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD2) = 1 / (62.5 * 1000000)
    
    TheHdw.Digital.Timing.ApplyTimingChanges True

    TheHdw.patterns(pathPathStr_freq).start

    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    
    TheHdw.Wait 0.1
    
    With TheHdw.Digital.pins(tx_diff).FreqCtr 'tx_diff
        .Clear
        .EventSlope = Positive
        .EventSource = vol
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With
    
    CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    CountVal = CountVal.Math.divide(TheHdw.Digital.pins(tx_diff).FreqCtr.Interval)
        
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)

    TheHdw.Digital.Patgen.HaltWait
    
    VihVal = TheHdw.Digital.pins("LPDPRX_RX_D0_P").Levels.value(chVih) ''GP_PCIE_RX0_P
    
'''    Theexec.Datalog.WriteComment "Programmed Vih : " + CStr(VihVal)
    
    Dim li As Long
    Dim vsite As Variant
    Dim pldPosMeas As New PinListData
    For li = 0 To 11
        If right(CountVal.pins(li).name, 2) = "_n" Then

        ElseIf right(CountVal.pins(li).name, 2) = "_p" Then
            pldPosMeas.AddPin CountVal.pins(li).name
            For Each vsite In TheExec.sites.Active
                pldPosMeas = CountVal.pins(li).value(vsite)
            Next vsite
        End If
                
    Next li
 '  judgeCheckResult pldPosMeas, 300, "LoopBack_62.5MHZ"
 Lo_Limit = 61500000
 Hi_Limit = 63500000
 Tname = "LoopBack_62.5MHz"
      TheExec.flow.TestLimit resultVal:=pldPosMeas, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="", ForceResults:=tlForceNone
      'glb_TestInstance = theexec.DataManager.instancename
      'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
   If TheExec.sites.ActiveCount = 0 Then Exit Function
'    theexec.Flow.TestLimit pldPosMeas, 50000000#, 75000000#, , , , , , "loopback_62.5MHZ", , , , , , , , 10000
    '***********************62.5MHZ end*******************************
    
    '***********************125MHZ start******************************
    TheHdw.Digital.Timing.period("DIBC_MeasFreq_2").value = 1 / (125 * 1000000)
    TheHdw.Digital.pins(rx_p).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD2) = 1 / (125 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD1) = 1 / (125 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD2) = 1 / (125 * 1000000)

    TheHdw.Digital.Timing.ApplyTimingChanges True

    TheHdw.patterns(pathPathStr_freq).start

    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    
    TheHdw.Wait 0.1
    
    With TheHdw.Digital.pins(tx_diff).FreqCtr
        .Clear
        .EventSlope = Positive
        .EventSource = vol
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With
    
    CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    CountVal = CountVal.Math.divide(TheHdw.Digital.pins("LPDPRX_RX_D_ODD").FreqCtr.Interval)
        
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)

    TheHdw.Digital.Patgen.HaltWait
    
    VihVal = TheHdw.Digital.pins("LPDPRX_RX_D0_P").Levels.value(chVih) ''GP_PCIE_RX0_P
    
'''    Theexec.Datalog.WriteComment "Programmed Vih : " + CStr(VihVal)
    
    Set pldPosMeas = Nothing 'clean data
    For li = 0 To 11
        If right(CountVal.pins(li).name, 2) = "_n" Then

        ElseIf right(CountVal.pins(li).name, 2) = "_p" Then
            pldPosMeas.AddPin CountVal.pins(li).name
            For Each vsite In TheExec.sites.Active
                pldPosMeas = CountVal.pins(li).value(vsite)
            Next vsite
        End If
                
    Next li
  ' judgeCheckResult pldPosMeas, 304, "LoopBack_125MHZ"
    Lo_Limit = 124000000
 Hi_Limit = 126000000
 Tname = "LoopBack_125MHz"
      TheExec.flow.TestLimit resultVal:=pldPosMeas, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="", ForceResults:=tlForceNone
      'glb_TestInstance = theexec.DataManager.instancename
      'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
   If TheExec.sites.ActiveCount = 0 Then Exit Function
'    theexec.Flow.TestLimit pldPosMeas, 100000000#, 150000000#, , , , , , "loopback_125MHZ"
    '***********************125MHZ end********************************
    
    '***********************250MHZ start******************************
    TheHdw.Digital.Timing.period("DIBC_MeasFreq_2").value = 1 / (200 * 1000000)
    TheHdw.Digital.pins(rx_p).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD2) = 1 / (200 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD1) = 1 / (200 * 1000000) / 2
    TheHdw.Digital.pins(rx_n).Timing.EdgeTime("DIBC_MeasFreq_2", chEdgeD2) = 1 / (200 * 1000000)
    
    TheHdw.Digital.Timing.ApplyTimingChanges True

    TheHdw.patterns(pathPathStr_freq).start

    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    
    TheHdw.Wait 0.2
    
    With TheHdw.Digital.pins(tx_diff).FreqCtr
        .Clear
        .EventSlope = Positive
        .EventSource = vol
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With
    
    CountVal = TheHdw.Digital.pins(tx_diff).FreqCtr.Read
    CountVal = CountVal.Math.divide(TheHdw.Digital.pins("LPDPRX_RX_D_ODD").FreqCtr.Interval)
        
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)

    TheHdw.Digital.Patgen.HaltWait
    
    VihVal = TheHdw.Digital.pins("LPDPRX_RX_D0_P").Levels.value(chVih) ''GP_PCIE_RX0_P
    
'''    Theexec.Datalog.WriteComment "Programmed Vih : " + CStr(VihVal)
    
    Set pldPosMeas = Nothing 'clean data
    For li = 0 To 11
        If right(CountVal.pins(li).name, 2) = "_n" Then

        ElseIf right(CountVal.pins(li).name, 2) = "_p" Then
            pldPosMeas.AddPin CountVal.pins(li).name
            For Each vsite In TheExec.sites.Active
                pldPosMeas = CountVal.pins(li).value(vsite)
            Next vsite
        End If
                
    Next li
   'judgeCheckResult pldPosMeas, 308, "LoopBack_200MHZ"
       Lo_Limit = 198000000
 Hi_Limit = 202000000
 Tname = "LoopBack_200MHz"
      TheExec.flow.TestLimit resultVal:=pldPosMeas, lowVal:=Lo_Limit, hiVal:=Hi_Limit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.4f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="", ForceResults:=tlForceNone
      'glb_TestInstance = theexec.DataManager.instancename
      'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
   If TheExec.sites.ActiveCount = 0 Then Exit Function
'    theexec.Flow.TestLimit pldPosMeas, 160000000#, 240000000#, , , , , , "loopback_250MHZ"
    '***********************250MHZ end********************************
    TheHdw.Digital.pins("LPDPRX_RX_D").Disconnect
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Read_Status_Reg(Validating_ As Boolean)
Dim i As Integer
Dim pat_modify(7) As String
Dim numcap As Long
Dim PinData As New PinListData
Dim regdata() As Long
Dim rawdata As String
Dim modify_data As String

    If Validating_ Then
        Call PrLoadPattern(".\Pattern\DIB\SPIROMReadStatus.PAT")
        Exit Function    ' Exit after validation
    End If

modify_data = ""
rawdata = ""
     TheHdw.Utility.pins("K02").State = tlUtilBitOn
     TheHdw.Utility.pins("K04").State = tlUtilBitOn
TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
Call SetupHRam
'TheHdw.Digital.Patterns.pat(".\Pattern\DIB\SPIROMReadStatus.PAT").Load
'TheExec.Datalog.WriteComment ("Read Status Reg")'comment by rita
get_status_reg2 &H5, 8

'''''get_status_reg2 &H70, 8
'''''
'''''get_status_reg2 &H65, 8
'''''
'''''get_status_reg2 &HC8, 8
'''''
'''''get_status_reg2 &H85, 8
  TheHdw.Utility.pins("K02").State = tlUtilBitOff
     TheHdw.Utility.pins("K04").State = tlUtilBitOff

End Function


Public Function get_status_reg2(data As Integer, data_len As Integer) As String
Dim i As Integer
Dim pat_modify() As String
Dim numcap As Long
Dim PinData As New PinListData
Dim regdata() As Long
Dim rawdata As String
Dim modify_data As String
modify_data = ""
get_status_reg2 = ""
ReDim pat_modify(data_len - 1)
Call get_data(data, pat_modify, data_len)
For i = 0 To data_len - 1
modify_data = modify_data + pat_modify(i)
Next i
Call TheHdw.Digital.pins("SPI1_MOSI").patterns(".\DIBChecker\PATTERNs\SPIROMReadStatus.PAT").ModifyVectorBlockData("", 4, pat_modify)
'Call thehdw.Digital.Pins("SPI0_MISO").Patterns(GetDIBCheckerPath & "Patterns\" & "SPIROMReadStatus.PAT").ModifyVectorBlockData("", 4, pat_modify) 'by rita
TheHdw.Wait 0.05

TheHdw.patterns(".\DIBCHECKER\PATTERNs\SPIROMReadStatus.PAT").start
TheHdw.Digital.Patgen.HaltWait
numcap = 8 'update by tyrone on 3.24
'numcap = thehdw.Digital.HRAM.CapturedCycles
'TheExec.Datalog.WriteComment ("Capture Cycle =" & CStr(numcap)) 'comment by rita
PinData = TheHdw.Digital.pins("SPI1_MISO").hram.PinPF(0, , numcap)
'added by rita
Dim site As Variant 'by rita
Dim sdbRegValue1 As New SiteDouble 'by rita
Dim sdbRegValue2 As New SiteDouble 'by rita
Dim sdbRegValue3 As New SiteDouble 'by rita
Dim sdbRegValue4 As New SiteDouble 'by rita
Dim sdbRegValue5 As New SiteDouble 'by rita
Dim pldMeasureValue As New PinListData
pldMeasureValue.AddPin "SPI1_MISO"
'end of addedj by rita
For Each site In TheExec.sites.Active
regdata = PinData.pins(0).value
For i = 0 To numcap - 1
If i Mod 4 = 0 Then
get_status_reg2 = get_status_reg2 + " "
End If


If regdata(i) = 2 Then
 get_status_reg2 = get_status_reg2 + "1"
Else
 get_status_reg2 = get_status_reg2 + "0"
End If
Next i
'comment by rita
'TheExec.Datalog.WriteComment ("Reg " + CStr(Hex(data)) + " , Binary Format: " + modify_data + ", Reg Value = " + get_status_reg2)
'TheExec.Datalog.WriteComment ("")

'added by rita
Dim ReadStatusReg As Integer
Dim ReadFlagStatusReg As Integer
Dim ReadEnhancedVolatileConfReg As Integer
Dim ReadExtendedAddressReg As Integer
Dim ReadVolatileConfReg As Integer

If (data = &H5) Then
ReadStatusReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
sdbRegValue1.value = ReadStatusReg
ElseIf (data = &H70) Then
ReadFlagStatusReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
sdbRegValue2.value = ReadFlagStatusReg
ElseIf (data = &H65) Then
ReadEnhancedVolatileConfReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
sdbRegValue3.value = ReadEnhancedVolatileConfReg
ElseIf (data = &HC8) Then
ReadExtendedAddressReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
sdbRegValue4.value = ReadExtendedAddressReg
ElseIf (data = &H85) Then
ReadVolatileConfReg = (val(regdata(0)) - 1) * 2 ^ 7 + (val(regdata(1)) - 1) * 2 ^ 6 + (val(regdata(2)) - 1) * 2 ^ 5 + (val(regdata(3)) - 1) * 2 ^ 4 + (val(regdata(4)) - 1) * 2 ^ 3 + (val(regdata(5)) - 1) * 2 ^ 2 + (val(regdata(6)) - 1) * 2 ^ 1 + (val(regdata(7)) - 1) * 2 ^ 0
sdbRegValue5.value = ReadVolatileConfReg
End If

'end of added by rita

Next site

'added by rita
If (data = &H5) Then
pldMeasureValue.pins("SPI1_MISO") = sdbRegValue1
'judgeCheckResult pldMeasureValue, 7005, "ReadStatusReg"
 TheExec.flow.TestLimit resultVal:=pldMeasureValue, lowVal:=0, hiVal:=0, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:="ReadStatusReg", ForceVal:=0, ForceUnit:=unitCustom, ForceResults:=tlForceNone
 'glb_TestInstance = theexec.DataManager.instancename
 'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
'''ElseIf (data = &H70) Then   'no need test
'''pldMeasureValue.Pins("SPI0_MISO") = sdbRegValue2
'''judgeCheckResult pldMeasureValue, 7006, "ReadFlagStatusReg"
'''ElseIf (data = &H65) Then
'''pldMeasureValue.Pins("SPI0_MISO") = sdbRegValue3
'''judgeCheckResult pldMeasureValue, 7007, "ReadEnhancedVolatileConfReg"
'''ElseIf (data = &HC8) Then
'''pldMeasureValue.Pins("SPI0_MISO") = sdbRegValue4
'''judgeCheckResult pldMeasureValue, 7008, "ReadExtendedAddressReg"
'''ElseIf (data = &H85) Then
'''pldMeasureValue.Pins("SPI0_MISO") = sdbRegValue5
'''judgeCheckResult pldMeasureValue, 7009, "ReadVolatileConfReg"
End If
'end of added by rita

End Function
Public Function Relay_Parasitic_Cap(PinName As String, Test_Name As String, Force_Curr As Double, First_RLY As String, Sec_RLY As String, Cap_Low As Double, Cap_Hi As Double, Para_Cap_Lo As Double, Para_Cap_Hi As Double, Validating_ As Boolean)

        TheHdw.Digital.pins(PinName).Disconnect '20200401
        
        TheHdw.Utility.pins(First_RLY).State = tlUtilBitOn
        TheHdw.Utility.pins(Sec_RLY).State = tlUtilBitOn
        '***connect PPMU instrument***'
        TheHdw.PPMU.pins(PinName).Connect
        '***-----------------------***'

        '********************capacitor main test begin********************'
        Dim strHSD_Pins As String
        Dim vsite As Variant
        Dim dblPeriod As Double
        Dim pldMeasureValue As New PinListData
        Dim dblTestMeasureCurrent As Double
        Dim dblTestChargeVoltage As Double
        Dim dblTestOffsetVoltage As Double
        Dim pldFailCount As New PinListData
        Dim strPatternStartLabel As String
        Dim strTestPattern As String
        Dim strPatternPath As String
        Dim Tname As String

        strHSD_Pins = PinName
        dblTestMeasureCurrent = 0.00001
        dblTestChargeVoltage = 0.8
        dblTestOffsetVoltage = 0.3

        dblPeriod = 20 * ns
        strPatternStartLabel = "CX_100usPer020ns"

        strTestPattern = "CX_HSD_" & strHSD_Pins & ".pat" '_PCIE_RCAL_N.pat"
        strPatternPath = ".\Pattern\DIB\"
        
        If Validating_ Then
            Call PrLoadPattern(strPatternPath & strTestPattern)
            Exit Function    ' Exit after validation
        End If

        With TheHdw.PPMU.pins(strHSD_Pins)
                .ForceV 0
                .Gate = tlOn
                TheHdw.Wait 0.005
                .Gate = tlOff
                TheHdw.Wait 0.002
                .Disconnect
        End With

        TheHdw.Digital.ApplyLevelsTiming False, True, True, , , , strHSD_Pins

        TheHdw.Digital.ConnectPins (strHSD_Pins)
        TheHdw.Wait 0.002
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '*****measure parasitic capacitor begin*****'

        Dim pldMeasureValueParasitic As New PinListData
        Dim dblOffsetVoltageParasitic As Double
        Dim dblChargeVoltageParasitic As Double
        Dim dblMeasureCurrentParasitic As Double
        Dim dblPeriodParasitic As Double
        Dim strPatternStartLabelParasitic As String
        
        TheHdw.Utility.pins(First_RLY).State = tlUtilBitOff
        TheHdw.Wait 0.003

        pldMeasureValueParasitic.AddPin (strHSD_Pins)
        For Each vsite In TheExec.sites.Active
                pldMeasureValueParasitic.pins(strHSD_Pins).value(vsite) = 0
        Next vsite

        dblMeasureCurrentParasitic = dblTestMeasureCurrent
        dblChargeVoltageParasitic = dblTestChargeVoltage
        dblOffsetVoltageParasitic = dblTestOffsetVoltage
        dblPeriodParasitic = dblPeriod
        strPatternStartLabelParasitic = strPatternStartLabel

        With TheHdw.PinLevels.pins(strHSD_Pins)
                .ModifyLevel chVih, 2
                .ModifyLevel chVil, -0.2
                .ModifyLevel chVol, dblOffsetVoltageParasitic
                .ModifyLevel chVch, 5.5
                .ModifyLevel chVcl, -1
        End With

        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVoh, dblOffsetVoltageParasitic + dblChargeVoltageParasitic

        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIoh, 0
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIol, dblMeasureCurrentParasitic
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVt, 5.5

        'If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
'                TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Load
       ' End If
        
        Dim iI As Integer
        Dim pldFailCountTemp As New PinListData
        
'''        TheHdw.Digital.Patterns.Pat(strPatternPath & strTestPattern).Run "PCIE_RCAL_N_" & strPatternStartLabelParasitic
'''        TheHdw.Wait 0.005
'''
'''        pldFailCount = TheHdw.Digital.Pins(strHSD_Pins).FailCount
'''        TheHdw.Wait 0.005

        For iI = 1 To 10
           TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Run strHSD_Pins & "_" & strPatternStartLabelParasitic
           TheHdw.Wait 0.005
           If (iI = 1) Then
               pldFailCount = TheHdw.Digital.pins(strHSD_Pins).FailCount
           Else
                pldFailCountTemp = TheHdw.Digital.pins(strHSD_Pins).FailCount
                pldFailCount = pldFailCount.Math.Add(pldFailCountTemp)
           End If
           TheHdw.Wait 0.005
        Next
        
        pldFailCount = pldFailCount.Math.divide(10)


        For Each vsite In TheExec.sites.Active
                pldMeasureValueParasitic.pins(strHSD_Pins).value(vsite) = dblMeasureCurrentParasitic * ((pldFailCount.pins(strHSD_Pins).value(vsite) * dblPeriodParasitic) / dblChargeVoltageParasitic)
        Next vsite

        If TheExec.TesterMode = testModeOffline Then
                For Each vsite In TheExec.sites.Active
                        pldMeasureValueParasitic.pins(strHSD_Pins).value = 0
                Next vsite
        End If

         TheHdw.Digital.DisconnectPins (strHSD_Pins)
        With TheHdw.PPMU.pins(strHSD_Pins)
                .ForceV 0
                .Connect
                .Gate = tlOn
                TheHdw.Wait 0.005
                .Gate = tlOff
                .Disconnect
                TheHdw.Wait 0.002
        End With
        
        TheHdw.Utility.pins(First_RLY).State = tlUtilBitOn
        
        TheHdw.Wait 0.003
        
        TheHdw.Digital.ConnectPins (strHSD_Pins)
        TheHdw.Wait 0.002
        '*****measure parasitic capacitor end*****'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       Dim pldMeasureValueParasitic1 As New PinListData
       Dim dblOffsetVoltageParasitic1 As Double
       Dim dblChargeVoltageParasitic1 As Double
       Dim dblMeasureCurrentParasitic1 As Double
       Dim dblPeriodParasitic1 As Double
       Dim strPatternStartLabelParasitic1 As String
    
        TheHdw.Utility.pins(Sec_RLY).State = tlUtilBitOff
       TheHdw.Wait 0.003
    
       pldMeasureValueParasitic1.AddPin (strHSD_Pins)
       For Each vsite In TheExec.sites.Active
          pldMeasureValueParasitic1.pins(strHSD_Pins).value(vsite) = 0
       Next vsite
    
       dblMeasureCurrentParasitic1 = dblTestMeasureCurrent
       dblChargeVoltageParasitic1 = dblTestChargeVoltage
       dblOffsetVoltageParasitic1 = dblTestOffsetVoltage
       dblPeriodParasitic1 = dblPeriod
       strPatternStartLabelParasitic1 = strPatternStartLabel
    
       With TheHdw.PinLevels.pins(strHSD_Pins)
          .ModifyLevel chVih, 2
          .ModifyLevel chVil, -0.2
          .ModifyLevel chVol, dblOffsetVoltageParasitic1
          .ModifyLevel chVch, 5.5
          .ModifyLevel chVcl, -1
       End With
    
       TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVoh, dblOffsetVoltageParasitic1 + dblChargeVoltageParasitic1
    
       TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIoh, 0
       TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIol, dblMeasureCurrentParasitic1
       TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVt, 5.5
    
       If CInt(left(TheExec.SoftwareVersion, 1)) < 8 Then
          TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Load
       End If
    '   TheHdw.Digital.Patterns.Pat(strPatternPath & strTestPattern).Run "PCIE_RCAL_N_" & strPatternStartLabelParasitic1
    '   TheHdw.Wait 0.005
    '
    '   pldFailCount = TheHdw.Digital.Pins(strHSD_Pins).FailCount
    '   TheHdw.Wait 0.005
       For iI = 1 To 10
               TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Run strHSD_Pins & "_" & strPatternStartLabelParasitic1
               TheHdw.Wait 0.005
               If (iI = 1) Then
                   pldFailCount = TheHdw.Digital.pins(strHSD_Pins).FailCount
               Else
                    pldFailCountTemp = TheHdw.Digital.pins(strHSD_Pins).FailCount
                    pldFailCount = pldFailCount.Math.Add(pldFailCountTemp)
               End If
               TheHdw.Wait 0.005
          Next
          pldFailCount = pldFailCount.Math.divide(10)
    
       For Each vsite In TheExec.sites.Active
          pldMeasureValueParasitic1.pins(strHSD_Pins).value(vsite) = dblMeasureCurrentParasitic1 * ((pldFailCount.pins(strHSD_Pins).value(vsite) * dblPeriodParasitic1) / dblChargeVoltageParasitic1)
       Next vsite
    
       If TheExec.TesterMode = testModeOffline Then
          For Each vsite In TheExec.sites.Active
             pldMeasureValueParasitic1.pins(strHSD_Pins).value = 0
          Next vsite
       End If
    
        TheHdw.Digital.DisconnectPins (strHSD_Pins)
       With TheHdw.PPMU.pins(strHSD_Pins)
          .ForceV 0
          .Connect
          .Gate = tlOn
          TheHdw.Wait 0.005
          .Gate = tlOff
          .Disconnect
          TheHdw.Wait 0.002
       End With
       TheHdw.Utility.pins(Sec_RLY).State = tlUtilBitOn
       TheHdw.Wait 0.003
    
       TheHdw.Digital.ConnectPins (strHSD_Pins)
       TheHdw.Wait 0.002
       '*****measure parasitic capacitor end*****'
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With TheHdw.PinLevels.pins(strHSD_Pins)
                .ModifyLevel chVih, 2
                .ModifyLevel chVil, -0.2
                .ModifyLevel chVol, dblTestOffsetVoltage
                .ModifyLevel chVch, 5.5
                .ModifyLevel chVcl, -1
        End With

        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVoh, dblTestOffsetVoltage + dblTestChargeVoltage

        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIoh, 0
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIol, dblTestMeasureCurrent
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVt, 5.5

        TheHdw.Wait 0.002

        If CInt(left(TheExec.SoftwareVersion, 1)) < 8 Then
                TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Load
        End If
        
        
'''        TheHdw.Digital.Patterns.Pat(strPatternPath & strTestPattern).Run "PCIE_RCAL_N_" & strPatternStartLabel
'''        TheHdw.Wait 0.005
'''
'''        pldMeasureValue.AddPin (strHSD_Pins)
'''        pldFailCount = TheHdw.Digital.Pins(strHSD_Pins).FailCount
'''        TheHdw.Wait 0.005

        For iI = 1 To 10
             TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Run strHSD_Pins & "_" & strPatternStartLabel
             TheHdw.Wait 0.005
             If (iI = 1) Then
                 pldFailCount = TheHdw.Digital.pins(strHSD_Pins).FailCount
             Else
                  pldFailCountTemp = TheHdw.Digital.pins(strHSD_Pins).FailCount
                  pldFailCount = pldFailCount.Math.Add(pldFailCountTemp)
             End If
             TheHdw.Wait 0.005
        Next
        pldFailCount = pldFailCount.Math.divide(10)
      
        pldMeasureValue.AddPin (strHSD_Pins)
        For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(strHSD_Pins).value(vsite) = dblTestMeasureCurrent * ((pldFailCount.pins(strHSD_Pins).value(vsite) * dblPeriod) / dblTestChargeVoltage)
        Next vsite

        With TheHdw.PPMU.pins(strHSD_Pins)
                .ForceV 0
                .Gate = tlOff
                TheHdw.Wait 0.002
        End With

        TheHdw.Digital.DisconnectPins (strHSD_Pins)

        pldMeasureValue = pldMeasureValue.Math.Subtract(pldMeasureValueParasitic)

        '********************main test begin********************'
       ' judgeCheckResult pldMeasureValue, 4700, "CL1, K73/K74 ON"
       Tname = Test_Name & "_Cap"
        TheExec.flow.TestLimit resultVal:=pldMeasureValue, PinName:=strHSD_Pins, lowVal:=Cap_Low, hiVal:=Cap_Hi, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************main test end********************'

        '********************subtest begin 1********************'
        'judgeCheckResult pldMeasureValueParasitic, 4701, "K73 OFF"
         Tname = First_RLY & " OFF Para Cap"
        TheExec.flow.TestLimit resultVal:=pldMeasureValueParasitic, PinName:=strHSD_Pins, lowVal:=Para_Cap_Lo, hiVal:=Para_Cap_Hi, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************subtest end 1********************'

        '********************subtest begin 2********************'
       ' judgeCheckResult pldMeasureValueParasitic1, 4702, "K74 OFF"
        Tname = Sec_RLY & " OFF Para Cap"
        TheExec.flow.TestLimit resultVal:=pldMeasureValueParasitic1, PinName:=strHSD_Pins, lowVal:=Para_Cap_Lo, hiVal:=Para_Cap_Hi, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        If TheExec.sites.ActiveCount = 0 Then Exit Function

        '********************CAPACITOR subtest end 2********************'

        '***disconnect PPMU instrument***'
        TheHdw.PPMU.pins(strHSD_Pins).Gate = tlOff
        TheHdw.Wait 0.002
        TheHdw.PPMU.pins(strHSD_Pins).Reset tlResetSettings
        TheHdw.PPMU.pins(strHSD_Pins).Disconnect
        '***-----------------------***'
        'Call SmartRelaySwitch("")
        
End Function


Public Function DCVI_Cap_meas(Test_PinName As String, Test_Name As String, C_LowLimit As Double, C_HiLimit As Double, Para_Cap_Lo As Double, Para_Cap_Hi As Double, dblTestMeasureCurrent As Double, Optional Relay_On As PinList)

       ' Call SmartRelaySwitch("")
       
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
        '***connect DCVI instrument***'
        TheHdw.DCVI.pins(Test_PinName).Connect tlDCVIConnectDefault
        TheHdw.DCVI.pins(Test_PinName).Disconnect tlDCVIConnectHighForce
        TheHdw.DCVI.pins(Test_PinName).LocalKelvin = True
        TheHdw.Wait 0.0005
        '***-----------------------***'

        '********************capacitor test begin********************'
        Dim vsite As Variant
        Dim intI As Integer
        Dim intJ As Integer
        Dim pldMeasureValue As New PinListData
        Dim v1 As New PinListData
        Dim v2 As New PinListData
        Dim deltaV As New PinListData
        Dim t1 As New PinListData
        Dim t2 As New PinListData
        Dim deltaT As New PinListData
        Dim dT As Double
        Dim DCVIResultDSP As New DSPWave
        Dim strDCVI_PinArray() As String
        Dim dblCapSampleRate As Double
        Dim dblCapSampleSize As Double
        Dim dblCurrentRange As Double
        Dim dblCaptureWait As Double
        Dim strTestDCVI_Pin As String
'        Dim dblTestMeasureCurrent As Double 'William 20200512
        Dim dblTestChargeVoltage As Double
        Dim dblTestOffsetVoltage As Double
        Dim iIndex1 As New SiteLong
        Dim iIndex2 As New SiteLong
        Dim Tname As String
    
        strTestDCVI_Pin = Test_PinName
'        dblTestMeasureCurrent = 0.001 ''0.0005
        dblTestChargeVoltage = 1.5  ''1.5
        dblTestOffsetVoltage = 0.8 ''0.8

        strDCVI_PinArray = Split(strTestDCVI_Pin, ",")

        dblCapSampleRate = 100 * kHz
        dblCapSampleSize = 512
        dblCaptureWait = dblCapSampleSize / dblCapSampleRate
        dblCurrentRange = 0.005

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmOff

        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.Add ("CapSignal")
        With TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals("CapSignal")
                .SampleRate = 100000
                .SampleSize = 512
                .LoadSettings
        End With

        '***set DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Current = 0.005
                .Voltage = 5
                .VoltageRange.value = 7
                .CurrentRange.Autorange = True
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 7
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'

        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '*****measure parasitic capacitor begin*****'

        Dim dblMeasureCurrentParasiticCap As Double
        Dim dblCurrentRangeParasiticCap As Double
        Dim pldMeasureValueParasitic As New PinListData
        Dim pldV1_parasitic As New PinListData
        Dim pldV2_parasitic As New PinListData
        Dim pldDeltaV_parasitic As New PinListData
        Dim pldT1_parasitic As New PinListData
        Dim pldT2_parasitic As New PinListData
        Dim pldDeltaT_parasitic As New PinListData
        Dim pldDT_parasitic As Double
        Dim DCVIResultDSP_parasitic As New DSPWave

        TheHdw.Wait 0.003

        dblMeasureCurrentParasiticCap = 0.000001
        dblCurrentRangeParasiticCap = 0.000001

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblMeasureCurrentParasiticCap
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait 0.00512
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'computing of the parasitic cap value
        For intJ = 0 To UBound(strDCVI_PinArray)
                pldV1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldV2_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT2_parasitic.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
'''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
'''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
'''''                                        Call DCVIResultDSP_parasitic.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
'''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP_parasitic = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

        'DCVIResultDSP_parasitic.Plot

                                iIndex1 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                pldV1_parasitic.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic.data(iIndex1)
                                pldV2_parasitic.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic.data(iIndex2)
                                pldDT_parasitic = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value - pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value
                                If pldDT_parasitic = 0 Then
                                        pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        pldDeltaV_parasitic.AddPin (strTestDCVI_Pin)
        pldDeltaV_parasitic = pldV2_parasitic.Math.Subtract(pldV1_parasitic)
        pldDeltaT_parasitic.AddPin (strTestDCVI_Pin)
        pldDeltaT_parasitic = pldT2_parasitic.Math.Subtract(pldT1_parasitic)

        pldMeasureValueParasitic.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValueParasitic = pldDeltaT_parasitic.Math.Multiply(dblMeasureCurrentParasiticCap)
        
        For Each site In TheExec.sites.Active
            If pldDeltaV_parasitic.Math.compare(EqualTo, 0) Then
                pldMeasureValueParasitic = 9999
            Else
                pldMeasureValueParasitic = pldMeasureValueParasitic.Math.divide(pldDeltaV_parasitic)
            End If
        Next site
 
       'pldMeasureValueParasitic = pldMeasureValueParasitic.Math.Divide(pldDeltaV_parasitic)

        
        For intJ = 0 To UBound(strDCVI_PinArray)
                If TheExec.TesterMode = testModeOffline Then
                        For Each vsite In TheExec.sites.Active
                                pldMeasureValueParasitic.pins(strDCVI_PinArray(intJ)).value = 0
                        Next vsite
                End If
        Next intJ

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.002
       ' Call SmartRelaySwitch("K96")
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
        TheHdw.Wait 0.003

'        TheHdw.DCVI.Pins("PAD_MTR_ANALOG_TEST_N").Disconnect tlDCVIConnectHighForce
        TheHdw.DCVI.pins(strTestDCVI_Pin).LocalKelvin = False
        TheHdw.DCVI.pins(strTestDCVI_Pin).Connect tlDCVIConnectDefault
        TheHdw.Wait 0.0005

        '*****measure parasitic capacitor end*****'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblTestMeasureCurrent
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait dblCaptureWait
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False

        'computing of the capacitor value
        For intJ = 0 To UBound(strDCVI_PinArray)
                v1.AddPin (strDCVI_PinArray(intJ))
                v2.AddPin (strDCVI_PinArray(intJ))
                t1.AddPin (strDCVI_PinArray(intJ))
                t2.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
''''                                        Call DCVIResultDSP.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

'        DCVIResultDSP.Plot

                                iIndex1 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                t1.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                t2.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                v1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(iIndex1)
                                v2.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(iIndex2)
                                dT = t2.pins(strDCVI_PinArray(intJ)).value - t1.pins(strDCVI_PinArray(intJ)).value
                                If dT = 0 Then
                                        t2.pins(strDCVI_PinArray(intJ)).value = t2.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        deltaV.AddPin (strTestDCVI_Pin)
        deltaV = v2.Math.Subtract(v1)
        deltaT.AddPin (strTestDCVI_Pin)
        deltaT = t2.Math.Subtract(t1)

        pldMeasureValue.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValue = deltaT.Math.Multiply(dblTestMeasureCurrent)
      
        For Each site In TheExec.sites.Active
            If deltaV.Math.compare(EqualTo, 0) Then
                pldMeasureValue = 9999
            Else
                pldMeasureValue = pldMeasureValue.Math.divide(deltaV)
            End If
        Next site
            
        'pldMeasureValue = pldMeasureValue.Math.Divide(deltaV)

        
        pldMeasureValue = pldMeasureValue.Math.Subtract(pldMeasureValueParasitic)

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmDefault
        
        '********************main test begin********************'
        'judgeCheckResult pldMeasureValue, 4714, "CV6|CV7|CV8, K96 ON"
        Tname = Test_Name & " Cap"
        TheExec.flow.TestLimit resultVal:=pldMeasureValue, PinName:=strTestDCVI_Pin, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************main test end********************'

        '********************subtest begin 1********************'
        
        'judgeCheckResult pldMeasureValueParasitic, 4715, "K96 OFF" ', 0, 0.000000005, "F"
        Tname = Relay_On & " OFF Para Cap"
        TheExec.flow.TestLimit resultVal:=pldMeasureValueParasitic, PinName:=strTestDCVI_Pin, lowVal:=Para_Cap_Lo, hiVal:=Para_Cap_Hi, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************subtest end 1********************'

        Set pldMeasureValue = Nothing
        '********************capacitor test end********************'

        '***disconnect DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .Gate = False
                TheHdw.Wait 0.002
                .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 0.0005
        '***-----------------------***'
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
       
End Function

Public Function DCVI_Cap_Meas_2(Test_PinName1 As String, Test_PinName2 As String, Test_Name As String, C_LowLimit As Double, C_HiLimit As Double, Para_Cap_Lo As Double, Para_Cap_Hi As Double, Optional Relay_On As PinList)

        Dim Pins_On() As String, Pin_Cnt_On As Long
        Dim All_TestPin As String
        'Call SmartRelaySwitch("")
        TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
        
        All_TestPin = Test_PinName1 & "," & Test_PinName2

        '***connect DCVI instrument***'
        TheHdw.DCVI.pins(All_TestPin).AlarmClear
        TheHdw.DCVI.pins(All_TestPin).Connect tlDCVIConnectDefault
        TheHdw.DCVI.pins(All_TestPin).Disconnect tlDCVIConnectHighSense
        TheHdw.DCVI.pins(All_TestPin).LocalKelvin = True

        TheHdw.Wait 0.0005
        '***-----------------------***'


        '***set DCVI instrument***'
        With TheHdw.DCVI.pins(Test_PinName2)
                .mode = tlDCVIModeVoltage
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Voltage = 0
                .Current = 0.02
                .VoltageRange.value = 0.05
                .CurrentRange.value = 0.02
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 0
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmDefault
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'


        '********************capacitor main test begin********************'
        Dim vsite As Variant
        Dim intI As Integer
        Dim intJ As Integer
        Dim pldMeasureValue As New PinListData
        Dim v1 As New PinListData
        Dim v2 As New PinListData
        Dim deltaV As New PinListData
        Dim t1 As New PinListData
        Dim t2 As New PinListData
        Dim deltaT As New PinListData
        Dim dT As Double
        Dim DCVIResultDSP As New DSPWave
        Dim strDCVI_PinArray() As String
        Dim dblCapSampleRate As Double
        Dim dblCapSampleSize As Double
        Dim dblCurrentRange As Double
        Dim dblCaptureWait As Double
        Dim strTestDCVI_Pin As String
        Dim dblTestMeasureCurrent As Double
        Dim dblTestChargeVoltage As Double
        Dim dblTestOffsetVoltage As Double
        Dim iIndex1 As New SiteLong
        Dim iIndex2 As New SiteLong

        strTestDCVI_Pin = Test_PinName1
        dblTestMeasureCurrent = 0.0002
        dblTestChargeVoltage = 0.4
        dblTestOffsetVoltage = 0.2

        strDCVI_PinArray = Split(strTestDCVI_Pin, ",")

        dblCapSampleRate = 100 * kHz
        dblCapSampleSize = 512
        dblCaptureWait = dblCapSampleSize / dblCapSampleRate
        dblCurrentRange = 0.005

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmOff

        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.Add ("CapSignal")
        With TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals("CapSignal")
                .SampleRate = 100000
                .SampleSize = 512
                .LoadSettings
        End With

        '***set DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Current = 0.005
                .Voltage = 5
                .VoltageRange.value = 7
                .CurrentRange.Autorange = True
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 7
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'

        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '*****measure parasitic capacitor begin*****'

        Dim dblMeasureCurrentParasiticCap As Double
        Dim dblCurrentRangeParasiticCap As Double
        Dim pldMeasureValueParasitic As New PinListData
        Dim pldV1_parasitic As New PinListData
        Dim pldV2_parasitic As New PinListData
        Dim pldDeltaV_parasitic As New PinListData
        Dim pldT1_parasitic As New PinListData
        Dim pldT2_parasitic As New PinListData
        Dim pldDeltaT_parasitic As New PinListData
        Dim pldDT_parasitic As Double
        Dim DCVIResultDSP_parasitic As New DSPWave
        
        TheHdw.Utility.pins(Pins_On(0)).State = tlUtilBitOff
        TheHdw.Wait 0.003

        dblMeasureCurrentParasiticCap = 0.000002
        dblCurrentRangeParasiticCap = 0.000002

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblMeasureCurrentParasiticCap
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait 0.00512
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'computing of the parasitic cap value
        For intJ = 0 To UBound(strDCVI_PinArray)
                pldV1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldV2_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT2_parasitic.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
''''                                        Call DCVIResultDSP_parasitic.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP_parasitic = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

'        DCVIResultDSP_parasitic.Plot

                                iIndex1 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                pldV1_parasitic.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic.data(iIndex1)
                                pldV2_parasitic.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic.data(iIndex2)
                                pldDT_parasitic = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value - pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value
                                If pldDT_parasitic = 0 Then
                                        pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        pldDeltaV_parasitic.AddPin (strTestDCVI_Pin)
        pldDeltaV_parasitic = pldV2_parasitic.Math.Subtract(pldV1_parasitic)
        pldDeltaT_parasitic.AddPin (strTestDCVI_Pin)
        pldDeltaT_parasitic = pldT2_parasitic.Math.Subtract(pldT1_parasitic)

        pldMeasureValueParasitic.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValueParasitic = pldDeltaT_parasitic.Math.Multiply(dblMeasureCurrentParasiticCap)
        pldMeasureValueParasitic = pldMeasureValueParasitic.Math.divide(pldDeltaV_parasitic)

        For intJ = 0 To UBound(strDCVI_PinArray)
                If TheExec.TesterMode = testModeOffline Then
                        For Each vsite In TheExec.sites.Active
                                pldMeasureValueParasitic.pins(strDCVI_PinArray(intJ)).value = 0
                        Next vsite
                End If
        Next intJ

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.2
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.002
        
        TheHdw.Utility.pins(Pins_On(1)).State = tlUtilBitOff
        TheHdw.Wait 0.003

        '*****measure parasitic capacitor end*****'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '*****measure parasitic capacitor begin*****'

        Dim dblMeasureCurrentParasiticCap1 As Double
        Dim dblCurrentRangeParasiticCap1 As Double
        Dim pldMeasureValueParasitic1 As New PinListData
        Dim pldV1_parasitic1 As New PinListData
        Dim pldV2_parasitic1 As New PinListData
        Dim pldDeltaV_parasitic1 As New PinListData
        Dim pldT1_parasitic1 As New PinListData
        Dim pldT2_parasitic1 As New PinListData
        Dim pldDeltaT_parasitic1 As New PinListData
        Dim pldDT_parasitic1 As Double
        Dim DCVIResultDSP_parasitic1 As New DSPWave
        
        'strTestDCVI_Pin = Test_PinName2
   
        dblMeasureCurrentParasiticCap1 = 0.000002
        dblCurrentRangeParasiticCap1 = 0.000002

        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblMeasureCurrentParasiticCap1
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait 0.00512
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'computing of the parasitic cap value
        For intJ = 0 To UBound(strDCVI_PinArray)
                pldV1_parasitic1.AddPin (strDCVI_PinArray(intJ))
                pldV2_parasitic1.AddPin (strDCVI_PinArray(intJ))
                pldT1_parasitic1.AddPin (strDCVI_PinArray(intJ))
                pldT2_parasitic1.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
'''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
'''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
'''''                                        Call DCVIResultDSP_parasitic1.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
'''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP_parasitic1 = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

'        DCVIResultDSP_parasitic1.Plot

                                iIndex1 = DCVIResultDSP_parasitic1.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP_parasitic1.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                pldT1_parasitic1.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                pldT2_parasitic1.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                pldV1_parasitic1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic1.data(iIndex1)
                                pldV2_parasitic1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP_parasitic1.data(iIndex2)
                                pldDT_parasitic1 = pldT2_parasitic1.pins(strDCVI_PinArray(intJ)).value - pldT1_parasitic1.pins(strDCVI_PinArray(intJ)).value
                                If pldDT_parasitic1 = 0 Then
                                        pldT2_parasitic1.pins(strDCVI_PinArray(intJ)).value = pldT2_parasitic1.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        pldDeltaV_parasitic1.AddPin (strTestDCVI_Pin)
        pldDeltaV_parasitic1 = pldV2_parasitic1.Math.Subtract(pldV1_parasitic)
        pldDeltaT_parasitic1.AddPin (strTestDCVI_Pin)
        pldDeltaT_parasitic1 = pldT2_parasitic1.Math.Subtract(pldT1_parasitic)

        pldMeasureValueParasitic1.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValueParasitic1 = pldDeltaT_parasitic1.Math.Multiply(dblMeasureCurrentParasiticCap)
        pldMeasureValueParasitic1 = pldMeasureValueParasitic1.Math.divide(pldDeltaV_parasitic)

        For intJ = 0 To UBound(strDCVI_PinArray)
                If TheExec.TesterMode = testModeOffline Then
                        For Each vsite In TheExec.sites.Active
                                pldMeasureValueParasitic1.pins(strDCVI_PinArray(intJ)).value = 0
                        Next vsite
                End If
        Next intJ

        'set the instrument to drive 0V (discharge the instrument)
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeVoltage
                .Voltage = 0
                .Current = 0.02
                .Gate = True
        End With
        TheHdw.Wait 0.01
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.002
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
        TheHdw.Wait 0.003

        '***connect DCVI instrument***'
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P,PAD_MTR_VREF_N").Disconnect tlDCVIConnectHighSense
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P,PAD_MTR_VREF_N").LocalKelvin = False
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P,PAD_MTR_VREF_N").Connect tlDCVIConnectDefault
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P").Disconnect tlDCVIConnectHighSense
'        TheHdw.DCVI.Pins("PAD_MTR_VREF_P").LocalKelvin = True
'        TheHdw.Wait 0.0005
        '***-----------------------***'


        '*****measure parasitic capacitor end*****'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '***set DCVI instrument***'
        With TheHdw.DCVI.pins(Test_PinName2)
                .mode = tlDCVIModeVoltage
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Voltage = 0
                .Current = 0.02
                .VoltageRange.value = 0.05
                .CurrentRange.value = 0.02
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 0
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmDefault
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'
        
        
        strTestDCVI_Pin = Test_PinName1
        
        
        
        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblTestMeasureCurrent
                .Voltage = 4
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger

        'switch on the gate and start source the constant current
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        TheHdw.Wait dblCaptureWait
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False

        'computing of the capacitor value
        For intJ = 0 To UBound(strDCVI_PinArray)
                v1.AddPin (strDCVI_PinArray(intJ))
                v2.AddPin (strDCVI_PinArray(intJ))
                t1.AddPin (strDCVI_PinArray(intJ))
                t2.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
''''                                        m_strCapWavePath = GetDIBCheckerPath & "waves\"
''''                                        m_strCapWaveFile = "DCVI_CapSignal.txt"
''''                                        Call DCVIResultDSP.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
''''                                        dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

'        DCVIResultDSP.Plot

                                iIndex1 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage)
                                iIndex2 = DCVIResultDSP.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                t1.pins(strDCVI_PinArray(intJ)).value = iIndex1 / dblCapSampleRate
                                t2.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate
                                v1.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(iIndex1)
                                v2.pins(strDCVI_PinArray(intJ)).value = DCVIResultDSP.data(iIndex2)
                                dT = t2.pins(strDCVI_PinArray(intJ)).value - t1.pins(strDCVI_PinArray(intJ)).value
                                If dT = 0 Then
                                        t2.pins(strDCVI_PinArray(intJ)).value = t2.pins(strDCVI_PinArray(intJ)).value + 0.000000000001
                                End If

                        End If
                Next vsite
        Next intJ

        deltaV.AddPin (strTestDCVI_Pin)
        deltaV = v2.Math.Subtract(v1)
        deltaT.AddPin (strTestDCVI_Pin)
        deltaT = t2.Math.Subtract(t1)

        pldMeasureValue.AddPin (strTestDCVI_Pin)

        'computing of the capacitance value: C= I * dt / dU
        pldMeasureValue = deltaT.Math.Multiply(dblTestMeasureCurrent)
        pldMeasureValue = pldMeasureValue.Math.divide(deltaV)

        pldMeasureValue = pldMeasureValue.Math.Subtract(pldMeasureValueParasitic)

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmDefault
        Dim Tname As String

        '********************main test begin********************'
        'judgeCheckResult pldMeasureValue, 4716, "CV1|CV2, K97/K98 ON"
         Tname = Test_Name & " Cap"
         TheExec.flow.TestLimit resultVal:=pldMeasureValue, PinName:=Test_PinName1, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        
        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************main test end********************'

        '********************subtest begin 1********************'
        'judgeCheckResult pldMeasureValueParasitic, 4717, "K97 OFF" ', 0, 0.000000005, "F"
         Tname = Pins_On(0) & " OFF Para Cap"
         TheExec.flow.TestLimit resultVal:=pldMeasureValueParasitic, PinName:=Test_PinName1, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************subtest end 1********************'

        '********************subtest begin 2********************'
       ' judgeCheckResult pldMeasureValueParasitic1, 4718, "K98 OFF"
        Tname = Pins_On(1) & " OFF Para Cap"
         TheExec.flow.TestLimit resultVal:=pldMeasureValueParasitic1, PinName:=Test_PinName1, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        If TheExec.sites.ActiveCount = 0 Then Exit Function

        '********************CAPACITOR subtest end 2********************'
        Set pldMeasureValue = Nothing
        '***disconnect DCVI instrument***'
        With TheHdw.DCVI.pins(All_TestPin)
                .Gate = False
                TheHdw.Wait 0.002
                .Reset tlResetSettings + tlResetConnections
        End With
        TheHdw.Wait 0.0005
        '***-----------------------***'
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
End Function


Public Function Relay_Parasitic_Cap_2(Test_PinName1 As String, Test_PinName2 As String, Test_Name As String, C_LowLimit As Double, C_HiLimit As Double, Validating_ As Boolean, Optional Para_Cap_Lo As Double, Optional Para_Cap_Hi As Double, Optional Relay_On As PinList)

       ' Call SmartRelaySwitch("K91,K92,K93,K94")
       
        Dim Pins_On() As String, Pin_Cnt_On As Long
        'Dim All_TestPin As String
        'Call SmartRelaySwitch("")
        TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On
         TheHdw.Utility.pins(Relay_On).State = tlUtilBitOn
        '***connect PPMU instrument***'
        TheHdw.PPMU.pins(Test_PinName1).Connect
        TheHdw.PPMU.pins(Test_PinName2).Connect
        '***-----------------------***'


        '***set PPMU instrument***'
        TheHdw.PPMU.pins(Test_PinName2).ForceV 0, 2 * mA
        TheHdw.PPMU.pins(Test_PinName2).Gate = tlOn
        TheHdw.Wait 0.01
        '***-----------------------***'


        '********************capacitor test begin********************'
        Dim strHSD_Pins As String
        Dim vsite As Variant
        Dim dblPeriod As Double
        Dim pldMeasureValue As New PinListData
        Dim dblTestMeasureCurrent As Double
        Dim dblTestChargeVoltage As Double
        Dim dblTestOffsetVoltage As Double
        Dim pldFailCount As New PinListData
        Dim strPatternStartLabel As String
        Dim strTestPattern As String
        Dim strPatternPath As String

        strHSD_Pins = Test_PinName1
        dblTestMeasureCurrent = 0.0000032
        dblTestChargeVoltage = 0.4
        dblTestOffsetVoltage = 0.2
'        thehdw.Utility.Pins("K92,K94").State = tlUtilBitOff
        dblPeriod = 20 * ns
        strPatternStartLabel = "CX_100usPer020ns"

        strTestPattern = "CX_HSD_" & strHSD_Pins & ".pat"
        strPatternPath = ".\Pattern\DIB\"
        
        If Validating_ Then
            Call PrLoadPattern(strPatternPath & strTestPattern)
            Exit Function    ' Exit after validation
        End If

        With TheHdw.PPMU.pins(strHSD_Pins)
                .ForceV 0
                .Gate = tlOn
                TheHdw.Wait 0.005
                .Gate = tlOff
                TheHdw.Wait 0.002
                .Disconnect
        End With

        TheHdw.Digital.ApplyLevelsTiming False, True, True, , , , strHSD_Pins

        TheHdw.Digital.ConnectPins (strHSD_Pins)
        TheHdw.Wait 0.002
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '*****measure parasitic capacitor begin*****'

   Dim pldMeasureValueParasitic As New PinListData
   Dim dblOffsetVoltageParasitic As Double
   Dim dblChargeVoltageParasitic As Double
   Dim dblMeasureCurrentParasitic As Double
   Dim dblPeriodParasitic As Double
   Dim strPatternStartLabelParasitic As String
   
   TheHdw.Utility.pins(Pins_On(2)).State = tlUtilBitOff
   TheHdw.Wait 0.003

   pldMeasureValueParasitic.AddPin (strHSD_Pins)
   For Each vsite In TheExec.sites.Active
      pldMeasureValueParasitic.pins(strHSD_Pins).value(vsite) = 0
   Next vsite

   dblMeasureCurrentParasitic = dblTestMeasureCurrent
   dblChargeVoltageParasitic = dblTestChargeVoltage
   dblOffsetVoltageParasitic = dblTestOffsetVoltage
   dblPeriodParasitic = dblPeriod
   strPatternStartLabelParasitic = strPatternStartLabel

   With TheHdw.PinLevels.pins(strHSD_Pins)
      .ModifyLevel chVih, 1.5
      .ModifyLevel chVil, -0.2 '0
      .ModifyLevel chVol, dblOffsetVoltageParasitic
      .ModifyLevel chVch, 5.5
      .ModifyLevel chVcl, -1
   End With

   TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVoh, dblOffsetVoltageParasitic + dblChargeVoltageParasitic

   TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIoh, 0
   TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIol, dblMeasureCurrentParasitic
   TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVt, 5.5

   'If CInt(Left(TheExec.SoftwareVersion, 1)) < 8 Then
'      TheHdw.Digital.Patterns.pat(strPatternPath & strTestPattern).Load
   'End If
'   TheHdw.Digital.Patterns.Pat(strPatternPath & strTestPattern).Run "PAD_MTR_RREF_P_" & strPatternStartLabelParasitic
'   TheHdw.Wait 0.005
'
'   pldFailCount = TheHdw.Digital.Pins(strHSD_Pins).FailCount
'   TheHdw.Wait 0.005
Dim iI As Integer
   Dim pldFailCountTemp As New PinListData
For iI = 1 To 20
           TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Run strHSD_Pins & "_" & strPatternStartLabelParasitic
           TheHdw.Wait 0.005
           If (iI = 1) Then
               pldFailCount = TheHdw.Digital.pins(strHSD_Pins).FailCount
           Else
                pldFailCountTemp = TheHdw.Digital.pins(strHSD_Pins).FailCount
                pldFailCount = pldFailCount.Math.Add(pldFailCountTemp)
           End If

           TheHdw.Wait 0.005
      Next
      pldFailCount = pldFailCount.Math.divide(10)
      
   For Each vsite In TheExec.sites.Active
      pldMeasureValueParasitic.pins(strHSD_Pins).value(vsite) = dblMeasureCurrentParasitic * ((pldFailCount.pins(strHSD_Pins).value(vsite) * dblPeriodParasitic) / dblChargeVoltageParasitic)
   Next vsite

   If TheExec.TesterMode = testModeOffline Then
      For Each vsite In TheExec.sites.Active
         pldMeasureValueParasitic.pins(strHSD_Pins).value = 0
      Next vsite
   End If

    TheHdw.Digital.DisconnectPins (strHSD_Pins)
   With TheHdw.PPMU.pins(strHSD_Pins)
      .ForceV 0
      .Connect
      .Gate = tlOn
      TheHdw.Wait 0.005
      .Gate = tlOff
      .Disconnect
      TheHdw.Wait 0.002
   End With
   TheHdw.Utility.pins(Pins_On(2)).State = tlUtilBitOn
   TheHdw.Wait 0.003

   TheHdw.Digital.ConnectPins (strHSD_Pins)
   TheHdw.Wait 0.002
   '*****measure parasitic capacitor end*****'
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With TheHdw.PinLevels.pins(strHSD_Pins)
                .ModifyLevel chVih, 2
                .ModifyLevel chVil, -0.2 '0
                .ModifyLevel chVol, dblTestOffsetVoltage
                .ModifyLevel chVch, 5.5
                .ModifyLevel chVcl, -1
        End With

        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVoh, dblTestOffsetVoltage + dblTestChargeVoltage

        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIoh, 0
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chIol, dblTestMeasureCurrent
        TheHdw.PinLevels.pins(strHSD_Pins).ModifyLevel chVt, 5.5

        TheHdw.Digital.ConnectPins (strHSD_Pins)
        If CInt(left(TheExec.SoftwareVersion, 1)) < 8 Then
                TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Load
        End If
        TheHdw.Wait 0.002

        If CInt(left(TheExec.SoftwareVersion, 1)) < 8 Then
                TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Load
        End If
'        TheHdw.Digital.Patterns.Pat(strPatternPath & strTestPattern).Run "PAD_MTR_RREF_P_" & strPatternStartLabel
'        TheHdw.Wait 0.005
        For iI = 1 To 20
           TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Run strHSD_Pins & "_" & strPatternStartLabel
           TheHdw.Wait 0.005
           If (iI = 1) Then
               pldFailCount = TheHdw.Digital.pins(strHSD_Pins).FailCount
           Else
                pldFailCountTemp = TheHdw.Digital.pins(strHSD_Pins).FailCount
                pldFailCount = pldFailCount.Math.Add(pldFailCountTemp)
           End If

           TheHdw.Wait 0.005
      Next
      pldFailCount = pldFailCount.Math.divide(10)
        pldMeasureValue.AddPin (strHSD_Pins)
       
        For Each vsite In TheExec.sites.Active
                pldMeasureValue.pins(strHSD_Pins).value(vsite) = dblTestMeasureCurrent * ((pldFailCount.pins(strHSD_Pins).value(vsite) * dblPeriod) / dblTestChargeVoltage)
        Next vsite

        With TheHdw.PPMU.pins(strHSD_Pins)
                .ForceV 0
                .Gate = tlOff
                TheHdw.Wait 0.002
        End With

        TheHdw.Digital.DisconnectPins (strHSD_Pins)
        pldMeasureValue = pldMeasureValue.Math.Subtract(pldMeasureValueParasitic)
        
        '********************main test begin********************'
       ' judgeCheckResult pldMeasureValue, 4719, "C4, K91/K92/K93/K94 ON"
       ''''' Tname = Test_Name & "_Cap"
        TheExec.flow.TestLimit resultVal:=pldMeasureValue, PinName:=strHSD_Pins, lowVal:=C_LowLimit, hiVal:=C_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Test_Name & "_Cap", ForceVal:=0, ForceUnit:=unitVolt, customUnit:="F", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
        If TheExec.sites.ActiveCount = 0 Then Exit Function
        '********************main test end********************'

        If CInt(left(TheExec.SoftwareVersion, 1)) < 8 Then
                TheHdw.Digital.patterns.pat(strPatternPath & strTestPattern).Unload
        End If
        Set pldMeasureValue = Nothing
        '********************capacitor test end********************'

        '***disconnect PPMU instrument***'
        TheHdw.PPMU.pins(Test_PinName1).Gate = tlOff
        TheHdw.PPMU.pins(Test_PinName2).Gate = tlOff
        TheHdw.Wait 0.002
        TheHdw.PPMU.pins(Test_PinName1).Reset tlResetSettings
        TheHdw.PPMU.pins(Test_PinName2).Reset tlResetSettings
        TheHdw.PPMU.pins(Test_PinName1).Disconnect
        TheHdw.PPMU.pins(Test_PinName2).Disconnect
        '***-----------------------***'

End Function

Public Function PPMU_Relay_FIMV_rising(TestPinName As String, RelayName As String, SW_LowLimit As Double, SW_HiLimit As Double)
     Dim pinname_str As String
     'pinname_str = "MEM0_XI0_0"
     'LP5_DDR_RREF_IO_UNUSED
     pinname_str = TestPinName
     Dim PPMUMeasure As New PinListData
     Dim PPMUMeas() As New SiteDouble
     Dim TimeMeas() As New SiteDouble
     Dim SWTime As New SiteDouble
     Dim i As Integer
     Dim j As Integer
     Dim site As Variant
     Dim RStep As Integer
     Dim RStep_2 As Integer
     Dim Start_time As Double
     Dim End_time As Double
     Dim Swtich_DeltaT As New PinListData
     
     For Each site In TheExec.sites.Selected
      
          SWTime(site) = 0
      
    Next site
     
'mem0_cfsb
     ReDim PPMUMeas(100) As New SiteDouble
     ReDim TimeMeas(100) As New SiteDouble
     RStep = 20
     RStep_2 = 100
        If (pinname_str = "SPI1_SSIN") Or (pinname_str = "SPI1_SCLK") Or (pinname_str = "SPI1_MISO") Or (pinname_str = "SPI1_MOSI") Then
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOn
        Else
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOff
'           TheHdw.Utility.pins("K54").State = tlUtilBitOn
'           TheHdw.Utility.pins("K51").State = tlUtilBitOn
'           TheHdw.Utility.pins("K52").State = tlUtilBitOn
        End If
      'thehdw.Utility.pins(RelayName).State = tlUtilBitOff
      TheHdw.Digital.pins(pinname_str).Disconnect
'      TheHdw.Digital.pins(Replace(pinname_str, "_N", "_P")).Disconnect
    If (pinname_str = "RT_CLK32768") Or (pinname_str = "XI0") Or (pinname_str = "XO0") Then
        TheHdw.Digital.pins(pinname_str & "_PA").Disconnect
    End If
'''    With thehdw.PPMU.pins(pinname_str)
'''        .ForceV 0#
'''        .Connect
'''        .Gate = tlOn
'''        .Gate = tlOff
'''        .ForceI 0.00000001, 0.00002
'''            '.ForceI 0.0001, 0.00002
'''        .Gate = tlOn
'''
'''    End With
  For Each site In TheExec.sites.Selected
'      With TheHdw.PPMU.pins(pinname_str)
 '       .ForceV 0#
  '      .Connect
   '     .Gate = tlOn
    '    .Gate = tlOff
     '   .ForceI 0.00000001, 0.00002
            '.ForceI 0.0001, 0.00002
      '  .Gate = tlOn   20200407
    
    If (pinname_str = "RT_CLK32768") Or (pinname_str = "XI0") Or (pinname_str = "XO0") Then
        With TheHdw.PPMU.pins(pinname_str & "_PA")
            .Gate = tlOff
            .Connect
            .Gate = tlOn
            TheHdw.Wait 0.001
            .ForceV 0, 0
        End With '20200416
    End If
    
'    With TheHdw.PPMU.pins(Replace(pinname_str, "_N", "_P"))
'        .Gate = tlOff
'        .Connect
'        .Gate = tlOn
'        TheHdw.Wait 0.001
'        .ForceV 0, 0
'    End With '20200416
    
    With TheHdw.PPMU.pins(pinname_str)
        .Gate = tlOff
        .Connect
        .Gate = tlOn
        TheHdw.Wait 0.001
        .ForceV 0, 0
        TheHdw.Wait 0.01
        .ForceI 0.00000001, 0.00002
        .ClampVHi = 3
        
    End With '20200407
      

      
        For i = 0 To RStep
            'thehdw.Wait 0.000001
            'PPMUMeasure = thehdw.PPMU.Pins(pinname_str).Read(tlPPMUReadMeasurements, 20)   'normal measure
            Start_time = TheExec.Timer(0)
            PPMUMeas(i) = TheHdw.PPMU.pins(pinname_str).Read(tlPPMUReadMeasurements, 2)   'normal measure
            'End_time = TheExec.Timer(Start_time)
            TimeMeas(i) = TheExec.Timer(Start_time)
            'TheExec.Datalog.WriteComment "ElapsedTime Pat Site (" & Site & ")" + Format(End_time * 1000#, "##0.000") + " msec"
        Next i
        
        If (pinname_str = "SPI1_SSIN") Or (pinname_str = "SPI1_SCLK") Or (pinname_str = "SPI1_MISO") Or (pinname_str = "SPI1_MOSI") Then
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOff
        Else
'            TheHdw.Utility.pins("K54").State = tlUtilBitOff
'            TheHdw.Utility.pins("K52").State = tlUtilBitOff
'            TheHdw.Utility.pins("K51").State = tlUtilBitOff
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOn
        End If
     
        For i = RStep + 1 To RStep_2
            'thehdw.Wait 0.000001
            'PPMUMeasure = thehdw.PPMU.Pins(pinname_str).Read(tlPPMUReadMeasurements, 20)   'normal measure
            Start_time = TheExec.Timer(0)
       
            PPMUMeas(i) = TheHdw.PPMU.pins(pinname_str).Read(tlPPMUReadMeasurements, 2)   'normal measure
            TimeMeas(i) = TheExec.Timer(Start_time)
        Next i
        
        If (pinname_str = "SPI1_SSIN") Or (pinname_str = "SPI1_SCLK") Or (pinname_str = "SPI1_MISO") Or (pinname_str = "SPI1_MOSI") Then
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOn
        Else
           TheHdw.Utility.pins(RelayName).State = tlUtilBitOff
'           TheHdw.Utility.pins("K54").State = tlUtilBitOff
        End If
        
  Next site
  
'  Dim sub_value As Double
'    For Each Site In TheExec.sites.Selected
'        TheExec.Datalog.WriteComment "Start " & PPMUMeas(RStep)
'        For i = 0 To 100 ''UBound(PPMUMeas)
''            sub_value = PPMUMeas(i).Subtract(PPMUMeas(i - 1))
'            sub_value = PPMUMeas(i)
'            If sub_value <> 0 Then
'                TheExec.Datalog.WriteComment "Site" & Site & "  Index  : " & i & " Difference sub : " & sub_value
'            End If
'        Next i
'    Next Site

''''  For i = 0 To RStep_2
''''     'TheExec.Datalog.WriteComment PPMUMeas(i)
''''     TheExec.Flow.TestLimit resultVal:=PPMUMeas(i), ScaleType:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=i, ForceResults:=tlForceNone
''''  Next i
     Dim Catch_index As New SiteLong
    For Each site In TheExec.sites.Selected
      For i = RStep + 1 To RStep_2
        If PPMUMeas(i).Subtract(PPMUMeas(i - 1)) < -0.01 Or PPMUMeas(i).Subtract(PPMUMeas(i - 1)) > 0.07 Then
            Catch_index = i
'            TheExec.Datalog.WriteComment "Site" & Site & "  Index  : " & i & " Difference sub : " & PPMUMeas(i).Subtract(PPMUMeas(i - 1)) & " <--"
            Exit For
'        If PPMUMeas(i) < PPMUMeas(i - 1) Then
'            Catch_index = i
'            Exit For
'''''        Else
'''''          TheExec.Datalog.WriteComment "Site" & Site & "  Index  : " & i & " Difference sub : " & PPMUMeas(i).Subtract(PPMUMeas(i - 1))
        End If
      Next i
    Next site
    
    For Each site In TheExec.sites.Selected
      For j = RStep + 1 To Catch_index(site)
          SWTime = SWTime + TimeMeas(j)
      Next j
    Next site
     Swtich_DeltaT.AddPin (pinname_str)
    Swtich_DeltaT = SWTime
    
    Dim Tname As String
     Tname = RelayName & " Switch_Time"
    TheExec.flow.TestLimit resultVal:=Swtich_DeltaT, lowVal:=SW_LowLimit, hiVal:=SW_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="S", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    
    
       TheHdw.Utility.pins(RelayName).State = tlUtilBitOff


    With TheHdw.PPMU.pins(pinname_str)
        .ForceV 0#
        .Connect
        .Gate = tlOn
    End With
    If (pinname_str = "RT_CLK32768") Or (pinname_str = "XI0") Or (pinname_str = "XO0") Then
        With TheHdw.PPMU.pins(pinname_str & "_PA")
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With '20200416
    End If
End Function

Public Function DCVI_RLY_Rising_Meas(Test_PinName1 As String, Test_Name As String, SW_LowLimit As Double, SW_HiLimit As Double, Optional Switch_Type As String = "Default open", Optional Relay_On As PinList)

        Dim Pins_On() As String, Pin_Cnt_On As Long
        Dim All_TestPin As String
        'Call SmartRelaySwitch("")
        TheExec.DataManager.DecomposePinList Relay_On, Pins_On(), Pin_Cnt_On
        TheHdw.Utility.pins(Relay_On).State = tlUtilBitOff
         TheHdw.Utility.pins("K07").State = tlUtilBitOff
        If Relay_On = "K07" Then
          TheHdw.Utility.pins("K08").State = tlUtilBitOn
        End If
        
        'All_TestPin = Test_PinName1 & "," & Test_PinName2

        '***connect DCVI instrument***'
        TheHdw.DCVI.pins(Test_PinName1).AlarmClear
        TheHdw.DCVI.pins(Test_PinName1).Connect tlDCVIConnectDefault
        TheHdw.DCVI.pins(Test_PinName1).Disconnect tlDCVIConnectHighSense
        TheHdw.DCVI.pins(Test_PinName1).LocalKelvin = True

        TheHdw.Wait 0.0005
        '***-----------------------***'


        '********************capacitor main test begin********************'
        Dim vsite As Variant
        Dim intI As Integer
        Dim intJ As Integer
        Dim pldMeasureValue As New PinListData
        Dim v1 As New PinListData
        Dim v2 As New PinListData
        Dim deltaV As New PinListData
        Dim t1 As New PinListData
        Dim t2 As New PinListData
        Dim deltaT As New PinListData
        Dim dT As Double
        Dim DCVIResultDSP As New DSPWave
        Dim strDCVI_PinArray() As String
        Dim dblCapSampleRate As Double
        Dim dblCapSampleSize As Double
        Dim dblCurrentRange As Double
        Dim dblCaptureWait As Double
        Dim strTestDCVI_Pin As String
        Dim dblTestMeasureCurrent As Double
        Dim dblTestChargeVoltage As Double
        Dim dblTestOffsetVoltage As Double
        Dim iIndex1 As New SiteLong
        Dim iIndex2 As New SiteLong
        'Dim RLY_CapSiignal As String

        
        strTestDCVI_Pin = Test_PinName1
        dblTestMeasureCurrent = 0.0002


        strDCVI_PinArray = Split(strTestDCVI_Pin, ",")
        dblCapSampleRate = 1000 * kHz
        dblCapSampleSize = 512
        dblCaptureWait = dblCapSampleSize / dblCapSampleRate
        dblCurrentRange = 0.005

        TheHdw.DCVI.pins(strTestDCVI_Pin).Alarm(tlDCVIAlarmMode) = tlAlarmOff
        
        'RLY_CapSiignal = "CapSignal"
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.Add ("CapSignal")
        With TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals("CapSignal")
                .SampleRate = 1000000
                .SampleSize = 256
                .LoadSettings
        End With

        '***set DCVI instrument***'
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .ComplianceRange(tlDCVICompliancePositive).value = 7
                .ComplianceRange(tlDCVIComplianceNegative).value = 2
                TheHdw.Wait 0.01
                .Current = 0 '0.005
                .Voltage = 5
                .VoltageRange.value = 7
                .CurrentRange.Autorange = True
                .NominalBandwidth = 1003
                .Meter.mode = tlDCVIMeterVoltage
                .Meter.VoltageRange = 7
                .Meter.Filter.bypass = False
                .Meter.Filter.value = 10000
                .Alarm(tlDCVIAlarmMode) = tlAlarmOff
                .Gate = True
                TheHdw.Wait 0.01
        End With
        '***-----------------------***'

        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

        'set the instrument to drive 0V (discharge the instrument)
'''''        With thehdw.DCVI.pins(strTestDCVI_Pin)
'''''                .mode = tlDCVIModeVoltage
'''''                .Voltage = 0
'''''                .Current = 0.2
'''''                .Gate = True
'''''        End With
'''''        thehdw.Wait 0.01
'''''        thehdw.DCVI.pins(strTestDCVI_Pin).Gate = False
'''''        thehdw.Wait 0.005

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '*****measure parasitic capacitor begin*****'

        Dim dblMeasureCurrentParasiticCap As Double
        Dim dblCurrentRangeParasiticCap As Double
        Dim pldMeasureValueParasitic As New PinListData
        Dim pldV1_parasitic As New PinListData
        Dim pldV2_parasitic As New PinListData
        Dim pldDeltaV_parasitic As New PinListData
        Dim pldT1_parasitic As New PinListData
        Dim pldT2_parasitic As New PinListData
        Dim Swtich_DeltaT As New PinListData
        Dim pldDT_parasitic As Double
        Dim DCVIResultDSP_parasitic As New DSPWave
        Dim RLY_CapDSP As New DSPWave
        Dim i As Integer
        Dim Step2 As Integer
        Set Swtich_DeltaT = Nothing
       ' thehdw.Utility.Pins(Pins_On(0)).State = tlUtilBitOff
        TheHdw.Wait 0.003
        Dim Start_time As Double
        Dim End_time As Double
        dblMeasureCurrentParasiticCap = 0.000001
        'dblMeasureCurrentParasiticCap = 0.000002
        'dblCurrentRangeParasiticCap = 0.000002
        dblCurrentRangeParasiticCap = 0.000002
        'set the instrument to drive the constant current to charge the capacitor
        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Current = dblMeasureCurrentParasiticCap
               ' .CurrentRange = dblCurrentRangeParasiticCap
                .Voltage = 5
        End With
        TheHdw.Wait 0.01

        'start capture
        TheHdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals.item("CapSignal").Trigger


        'switch on the gate and start source the constant current
        
        Start_time = TheExec.Timer(0)
        
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = True
        'thehdw.Wait 0.0001
        End_time = TheExec.Timer(Start_time)
        TheHdw.Utility.pins(Pins_On(0)).State = tlUtilBitOn
        
        
        
        TheHdw.Wait 0.00512
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005
        'RLY_CapDSP = thehdw.DCVI.pins(strTestDCVI_Pin).Capture.Signals("CapSignal").DSPWave
        
       'TheExec.Datalog.WriteComment "ElapsedTime Pat Site (" & Site & ")" + Format(End_time * 1000#, "##0.000") + " msec"
        'computing of the parasitic cap value
        
        For intJ = 0 To UBound(strDCVI_PinArray)

                pldT1_parasitic.AddPin (strDCVI_PinArray(intJ))
                pldT2_parasitic.AddPin (strDCVI_PinArray(intJ))
                Swtich_DeltaT.AddPin (strDCVI_PinArray(intJ))

                For Each vsite In TheExec.sites.Active
                        If TheExec.sites.Active Then

                                If TheExec.TesterMode = testModeOffline Then
                                       ' m_strCapWavePath = GetDIBCheckerPath & "waves\"
                                       ' m_strCapWaveFile = "DCVI_CapSignal.txt"
                                       ' Call DCVIResultDSP_parasitic.FileImport(m_strCapWavePath & m_strCapWaveFile, File_txt)
                                       ' dblCapSampleSize = 500
                                Else
                                        DCVIResultDSP_parasitic = TheHdw.DCVI.pins(strDCVI_PinArray(intJ)).Capture.Signals.item("CapSignal").DSPWave.pins(strDCVI_PinArray(intJ)).value
                                End If

'        DCVIResultDSP_parasitic.Plot

                                iIndex1 = DCVIResultDSP_parasitic.FindIndex(OfFirstElement, GreaterThan, 0.01)
                                 
                                Step2 = iIndex1 + 1
                                For i = Step2 To 255
                                   If DCVIResultDSP_parasitic.Element(i) < DCVIResultDSP_parasitic.Element(i - 1) Then
                                       iIndex2 = i - 1
                                       Exit For
                                   End If
                                Next i
                                
                                'iIndex2 = DCVIResultDSP_parasitic.FindIndex(OfLastElement, LessThan, dblTestOffsetVoltage + dblTestChargeVoltage)
                                pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value = End_time
                                'pldT1_parasitic.Pins(strDCVI_PinArray(intJ)).Value = iIndex1 / dblCapSampleRate
                                pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value = iIndex2 / dblCapSampleRate

                                pldDT_parasitic = pldT2_parasitic.pins(strDCVI_PinArray(intJ)).value - pldT1_parasitic.pins(strDCVI_PinArray(intJ)).value
'                                If pldDT_parasitic = 0 Then
'                                        pldT2_parasitic.Pins(strDCVI_PinArray(intJ)).Value = pldT2_parasitic.Pins(strDCVI_PinArray(intJ)).Value + 0.000000000001
'                                End If
                              ' Swtich_DeltaT.AddPin (strDCVI_PinArray(intJ))
                               Swtich_DeltaT = pldDT_parasitic

                        End If
                Next vsite
        Next intJ
Dim Tname As String
Tname = Test_Name & " Switch_Time"
TheExec.flow.TestLimit resultVal:=Swtich_DeltaT, lowVal:=SW_LowLimit, hiVal:=SW_HiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="S", ForceResults:=tlForceNone
    'glb_TestInstance = theexec.DataManager.instancename
    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    TheHdw.Utility.pins(Pins_On(0)).State = tlUtilBitOff

        With TheHdw.DCVI.pins(strTestDCVI_Pin)
                .mode = tlDCVIModeCurrent
                .Voltage = 0
                .Current = 0
                .Gate = True
        End With
        TheHdw.Wait 0.05
        TheHdw.DCVI.pins(strTestDCVI_Pin).Gate = False
        TheHdw.Wait 0.005

End Function


'measure the delta of rising point when the relay ON and Off
'the delta will be same as TDR difference
Public Function measureRelayTDRDelta(PinName As String, RelayName As String, loLimit As Double, hiLimit As Double, Validating_ As Boolean) As Long
    'setup default parameters
    Dim RelayOffTime As New SiteDouble
    Dim RelayOnTime As New SiteDouble
    Dim RelayDeltaTime As New SiteDouble
    Dim sPinName As String
    Dim sRelayName As String
    
    If Validating_ Then
        Call PrLoadPattern(".\Pattern\DIB\DIBC_TDR_ALL__.pat")
        Exit Function    ' Exit after validation
    End If
    
    
    sPinName = PinName
    sRelayName = RelayName
    'Call ApplyDIBCheckerTimingLevel
    TheHdw.PPMU.pins(PinName).Disconnect
    
    TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, PinLevelsSheet:="DIBC_Levels__", TimeSetSheet:="DIBC_TSets__"
    'Call SetLevelForTDRCheck
''''    Dim strAllDigitalPins()  As String
''''    Dim lngPnum As Long
   
''''     Call TheExec.DataManager.GetPinNames(strAllDigitalPins, chIO, lngPnum)
''''     m_sAllDigitalPinList = Join(strAllDigitalPins, ",")
     '**********************************************
         If m_sAllDigitalPinList = vbNullString Then
        Dim strAllDigitalPins()  As String
        Dim lngPnum As Long
        Call TheExec.DataManager.GetPinNames(strAllDigitalPins, chIO, lngPnum)
        m_sAllDigitalPinList = Join(strAllDigitalPins, ",")
        End If
    If m_sAllDigitalPinList <> vbNullString Then
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVil) = 0.1  'm_sAllDigitalPinList
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVih) = 1.1
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVol) = 0
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVoh) = 0
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVt) = 0.1
        'thehdw.Digital.Pins(m_sAllDigitalPinList).Levels.Value(chIol) = 0
        'thehdw.Digital.Pins(m_sAllDigitalPinList).Levels.Value(chIoh) = 0
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVcl) = -1#
        TheHdw.Digital.pins(m_sAllDigitalPinList).Levels.value(chVch) = 6
    End If
     
     '**********************************************
    'Call ApplyDIBCheckerTDRPattern
'    TheHdw.Patterns(".\Pattern\DIB\DIBC_TDR_ALL__.pat").Load
    TheHdw.patterns(".\Pattern\DIB\DIBC_TDR_ALL__.pat").start
    TheHdw.Digital.pins(sPinName).Connect
    TheHdw.Wait 0.001
    ' Set the VOh value
    TheHdw.Digital.pins(sPinName).Levels.value(chVoh) = 0.8
                                                                                               
     If sPinName = "DDR2_RREF" Then ' for t4402 Kevin0601
         TheHdw.Digital.pins(sPinName).Levels.value(chVoh) = 0.79 '0.85 Kevin0601,0.8 Kevin0823
'     ElseIf sPinName = "DDR1_RREF" Then ' for t4401
'         thehdw.Digital.Pins(sPinName).Levels.Value(chVoh) = 0.85
     ElseIf sPinName = "DDR0_RREF" Or sPinName = "DDR3_RREF" Then ' for t4400,4403 Kevin0601
         TheHdw.Digital.pins(sPinName).Levels.value(chVoh) = 0.75
                                                                                               
     End If
     TheHdw.Wait 0.001
    ' Find the time with relay Off
    TheHdw.Utility.pins(sRelayName).State = tlUtilBitOff
    RelayOffTime = searchTDRRisingTime(sPinName)
                                                                                               
    ' Find the time with relay on
    TheHdw.Utility.pins(sRelayName).State = tlUtilBitOn
                                                                                               
    RelayOnTime = searchTDRRisingTime(sPinName)
    RelayDeltaTime = RelayOnTime.Subtract(RelayOffTime)
                                                                                               
    Dim plDeltaValue As New PinListData
    Dim thissite As Variant
    plDeltaValue.AddPin (sPinName)
    For Each thissite In TheExec.sites.Selected
        plDeltaValue.pins(sPinName).value = RelayDeltaTime.Abs
    Next thissite
    Dim Tname As String
    Tname = sRelayName & "  TDR"
    'Set measureRelayTDRDelta = plDeltaValue
    TheExec.flow.TestLimit resultVal:=plDeltaValue, lowVal:=loLimit, hiVal:=hiLimit, unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=0, ForceUnit:=unitNone, customUnit:="S", ForceResults:=tlForceNone
       'glb_TestInstance = theexec.DataManager.instancename
       ' theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    TheHdw.Utility.pins(sRelayName).State = tlUtilBitOff
     TheHdw.Digital.pins(sPinName).Disconnect
     TheHdw.Wait 0.001
    'Call DisconnectAllPins
End Function

Public Function searchTDRRisingTime(sPinName As String) As SiteDouble
    Dim StartTime As Double
    Dim StopTime As Double
    Dim StepSize As Double
    'setup default parameters
    StartTime = 20 * ns
    StopTime = 80 * ns
    StepSize = 0.1 * ns
    
    Dim sdRisingTime As New SiteDouble
    Dim timeval As Double
    Dim plFailCount As PinListData
    Dim thissite As Variant
    Dim foundtime As New SiteBoolean
              
    For Each thissite In TheExec.sites.Selected
        foundtime = False
        sdRisingTime = 0
    Next thissite

    timeval = StartTime
    Do While timeval < StopTime
        ' Set the Strobe time
       ' thehdw.Digital.pins(sPinName).Timing.EdgeTime(GetTDRTimingSetName, chEdgeR0) = timeval
         TheHdw.Digital.pins(sPinName).Timing.EdgeTime("DIBC_TDR__", chEdgeR0) = timeval
        ' Run the pattern
        TheHdw.Digital.Patgen.Restart
        TheHdw.Digital.Patgen.HaltWait
        Set plFailCount = TheHdw.Digital.pins(sPinName).FailCount
        For Each thissite In TheExec.sites.Selected
            ' Read the failcount
            If Not foundtime Then
                If plFailCount = 0 Then
                    foundtime = True
                    sdRisingTime = timeval
                End If
            End If
        Next thissite
        ' Found for all sites?
        If foundtime.All(True) Then
            Exit Do
        End If
        timeval = timeval + StepSize
    Loop
''    Dim Tname As String
''Tname = Test_Name & " TDR_Time"
''TheExec.Flow.TestLimit resultVal:=sdRisingTime, lowVal:=SW_LowLimit, hiVal:=SW_HiLimit, unit:=unitCustom, ScaleType:=scaleNone, formatStr:="%.3f", Tname:=Tname, forceVal:=0, ForceUnit:=unitNone, customUnit:="S", ForceResults:=tlForceNone

       
    Set searchTDRRisingTime = sdRisingTime
End Function
'''''Public Sub ErrorDescription(funcName As String)
''''''error description printing
'''''    Dim TestInstanceName As String
'''''    TestInstanceName = TheExec.DataManager.instanceName
'''''
'''''    TheExec.Datalog.WriteComment "TestInstance: " & TestInstanceName & ", " & funcName & " error, Err Code: " & err.number & ", Err Description: " & err.Description
'''''End Sub


'Public Function is_reference_installed(S As String) As Boolean
'    Dim x As Variant
'    is_reference_installed = False
'    For Each x In Application.ActiveWorkbook.VBProject.References
'        If S = x.name Then
'            is_reference_installed = True
'        End If
'    Next x
'End Function


'---------------------------------Test_Summary 20200413-------
Public Function Test_Summary() As Long
    Dim wb_def As Worksheet
    Dim wb As Workbook
    Dim Res_Row As Long, Res_Col As Long
    Dim Flag_Name As String
    Dim Flag_Name_List() As String
    Dim Test_Name As String
    Dim Test_Name_List() As String
    Dim site As Variant
    Dim Pass_Fail_Flag As New SiteBoolean
    Dim Fail_Flag As Boolean
    Dim i As Long
    Dim m_dlogstr As String
    Dim max_str As String
    Dim max_str_length As Long
    On Error GoTo errHandler
    
    TheExec.Datalog.WriteComment ("    --------------------------------Test Summary--------------------------------" + vbCrLf + "")
    Set wb = Application.ActiveWorkbook
    Set wb_def = wb.Sheets("DIB_Bin_Table")
    Res_Row = 4
    Res_Col = 3
    max_str_length = 0
    wb_def.Select
    While wb_def.Cells(Res_Row, Res_Col).value <> ""
        ReDim Preserve Flag_Name_List(Res_Row - 4)
        ReDim Preserve Test_Name_List(Res_Row - 4)
        Flag_Name = CStr(wb_def.Cells(Res_Row, Res_Col).value)
        Test_Name = Replace(Replace(Flag_Name, "F_", ""), "_Flag", "")
        Flag_Name_List(Res_Row - 4) = Flag_Name
        Test_Name_List(Res_Row - 4) = Test_Name
        Res_Row = Res_Row + 1
        
        If Len(Test_Name) > max_str_length Then
            max_str_length = Len(Test_Name)
            max_str = Test_Name & "    "
        End If
    Wend
    
'    For i = 0 To UBound(Flag_Name_List)
'        Fail_Flag = False
'        For Each Site In TheExec.sites
'            Pass_Fail_Flag(Site) = TheExec.Flow.SiteFlag(Site, Flag_Name_List(i))
'            If TheExec.Flow.SiteFlag(Site, Flag_Name_List(i)) = -1 Then
'                Pass_Fail_Flag(Site) = False
'            Else
'                Pass_Fail_Flag(Site) = TheExec.Flow.SiteFlag(Site, Flag_Name_List(i))
'            End If
'
'            If Pass_Fail_Flag(Site) = True Then
'                m_dlogstr = "Site(" + CStr(Site) + ") " + Test_Name_List(i) + " : Fail"
'                TheExec.Datalog.WriteComment (m_dlogstr)
'                Fail_Flag = True
'            End If
'
'        Next Site
'        If Fail_Flag = True Then
'            TheExec.Datalog.WriteComment ("")
'        End If
'    Next i
    
    Test_Name = max_str
    LSet Test_Name = "Test Instance"
    m_dlogstr = vbTab + "Site" + vbTab + Test_Name + vbTab + "Pass/Fail"
    TheExec.Datalog.WriteComment m_dlogstr
    TheExec.Datalog.WriteComment "    -----------------------------------------------------------------------"
    For Each site In TheExec.sites
        Fail_Flag = False
        For i = 0 To UBound(Flag_Name_List)
            Pass_Fail_Flag(site) = TheExec.flow.SiteFlag(site, Flag_Name_List(i))
            If TheExec.flow.SiteFlag(site, Flag_Name_List(i)) = -1 Then
                Pass_Fail_Flag(site) = False
            Else
                Pass_Fail_Flag(site) = TheExec.flow.SiteFlag(site, Flag_Name_List(i))
            End If
            
        
            If Pass_Fail_Flag(site) = True Then
                LSet Test_Name = Test_Name_List(i)
                m_dlogstr = vbTab + CStr(site) + vbTab + Test_Name + vbTab + "Fail"
                TheExec.Datalog.WriteComment m_dlogstr
                TheExec.AddOutput CStr(site) + Space(2) + Test_Name + Space(2) + "Fail", vbRed
                Fail_Flag = True
            End If
        Next i
        If Fail_Flag = False Then
            LSet Test_Name = ""
            TheExec.Datalog.WriteComment vbTab + CStr(site) + vbTab + Test_Name + vbTab + "Pass"
            TheExec.AddOutput CStr(site) + Space(2) + Test_Name + Space(2) + "Pass", vbBlack
        End If
        TheExec.Datalog.WriteComment ""
    Next site
'    m_dlogstr = vbTab + "Site(" + "0" + ") " + vbTab + Test_Name_List(0) + vbTab + " : Fail"
'    TheExec.Datalog.WriteComment m_dlogstr
'    TheExec.Datalog.WriteParametricResult (Test_Name_List(0))
    TheExec.Datalog.WriteComment ("    --------------------------------End Summary--------------------------------" + vbCrLf + "")
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "Error in Test Summary"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function DIB_RegKeyRead(i_RegKey As String)
    Dim myWS As Object
    Dim RegKeyRead As String
    
    On Error GoTo errHandler
    
    Set myWS = CreateObject("WScript.Shell")
    RegKeyRead = myWS.RegRead("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\" & i_RegKey)
    
    If RegKeyRead <> "" Then
        TheExec.flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=RegKeyRead, ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    Else
        TheExec.flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="No ProbeCard_ID", ForceResults:=tlForceNone
        'glb_TestInstance = theexec.DataManager.instancename
        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    End If
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "Error in DIB_RegKeyRead"
    If AbortTest Then Exit Function Else Resume Next

End Function
