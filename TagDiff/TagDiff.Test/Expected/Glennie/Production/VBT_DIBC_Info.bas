Attribute VB_Name = "VBT_DIBC_Info"
Public DIBC_F As Boolean

Public Function ExecuteIndex()

On Error GoTo errHandler

' TheExec.Datalog.WriteComment "Run Index: " & TheExec.ExecutionCount
   ' theexec.AddOutput theexec.ExecutionCount, vbBlue

'For AutoZ short TP
  '  theexec.EnableWord("Temp_25C") = True
'Flag


'

'    For Each log_StreamName In TheExec.Datalog.Setup.DataStreams.List
'        log_temp = Trim(TheExec.Datalog.Setup.DataStreams.Item(log_StreamName).Outputs.Text.OutputFile)
'
'            If log_temp <> "" Then
'                TheExec.AddOutput "log_StreamName:  " & log_StreamName, vbBlack
'                TheExec.AddOutput "log_temp:  " & log_temp, vbBlack
'                TheExec.AddOutput "_________", vbGreen
'            End If
'
     '--------------------------------------------------------
     'deleted
     
        'TheExec.Datalog.Setup.DataStreams.DeleteStream ("DIBC_Datalog_005")
        'TheExec.Datalog.ApplySetup
    '--------------------------------------------------------
    
    ''TheExec.Datalog.Setup.DatalogSetup.DatalogOn = True
    ''TheExec.Datalog.Setup.DataStreams.CreateStream ("DIBC_Datalog_001")
    ''TheExec.Datalog.Setup.DatalogSetup.SelectSetupFile("D:\Local_TP\MMM\DIBC_Leak.stp") = True
    ''TheExec.Datalog.Setup.DatalogSetup.TextOutputFile() = "D:\Local_TP\MMM\123.txt"
    '''DIBC = TheExec.Datalog.Setup.DatalogSetup.SetupFile()
    ''TheExec.Datalog.ApplySetup
'    Next


    'If TheExec.ExecutionCount = 0 Or TheExec.ExecutionCount < 50 And TheExec.ExecutionCount Mod 4 = 0 Or DIBC_F = True Then
    If DIBC_F = True Then
       ' theexec.AddOutput "DIBC is Runing....", vbGreen
        TheExec.Datalog.WriteComment "===DIBC TESTING Started==="
        
        'data stream
        
        ''Dim log_StreamName, log_tempA
        ''
        ''TheExec.Datalog.Setup.DatalogSetup.DatalogOn = True
        ''TheExec.Datalog.Setup.DataStreams.CreateStream ("DIBC_Datalog_005")
        ''
        ''With TheExec.Datalog.Setup.DataStreams("DIBC_Datalog_005")
        ''    .Outputs.Text.Enable
        ''    .Outputs.STDF.disable
        ''    .Outputs.Text.OverwriteOutput = True
        ''    .Outputs.Custom.Enable
        ''End With
        ''TheExec.Datalog.ApplySetup
        
        
    
        
             If TheExec.enableWord("DIBC_Enable") = True Then
        
               ' theexec.AddOutput "DIBC EnableWord is True", vbBlue
                
             Else
                'TheExec.AddOutput "DIBC EnableWord is False ,Enable It", vbRed
                TheExec.flow.enableWord("DIBC_Enable") = True
               ' theexec.AddOutput "DIBC EnableWord Enable", vbBlue
            
             End If
             
        
    Else
        
         TheExec.flow.enableWord("DIBC_Enable") = False
         'TheExec.Datalog.Setup.DataStreams.DeleteStream ("DIBC_Datalog_001")
        ' theexec.AddOutput "DIBC Flag is OFF ", vbRed
    End If



Exit Function
errHandler:
     TheExec.Datalog.WriteComment "error in ExecuteIndex"
    If AbortTest Then Exit Function Else Resume Next
         
End Function


