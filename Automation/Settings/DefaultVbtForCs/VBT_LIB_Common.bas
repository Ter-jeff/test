Attribute VB_Name = "VBT_LIB_Common"
Public Function GetBarcode()
    
'    Dim start_time As Double
'    Dim end_time As Double
'    Dim test_time As Double
'    start_time = TheExec.Timer
   
    Dim BarcodeStrInfor() As String
    Dim BarcodeHeader As String
    Dim i As Long
    Dim BarcodeSiteVariant As New SiteVariant
    Dim IEDAInputStr As String
    Dim SiteCount As Integer
    SiteCount = TheExec.Sites.Existing.Count
    ReDim IEDAInputStrArr(SiteCount - 1) As String
    'theexec.sites.Active
    If BarcodeString <> "" Then
        BarcodeHeader = Split(Replace(BarcodeString, ";", ""), ":")(0)
        BarcodeStrInfor = Split(Split(Replace(BarcodeString, ";", ""), ":")(1), ",")
        If BarcodeHeader <> "" Then
            
            '20240502 Mason - Clear Handler2DID Registry
            Call IEDA_Initialize(IEDAInputStr)
            
            TheExec.Datalog.WriteComment "---------Capture 2DID Information Start -------------"
            TheExec.Datalog.WriteComment "BarcodeString: " & BarcodeString
            For i = 0 To UBound(BarcodeStrInfor)
                If BarcodeStrInfor(i) <> "0" Then
                    TheExec.Datalog.WriteComment " Site = " & UBound(BarcodeStrInfor) - i & ", Barcode = " & BarcodeStrInfor(i)
                    BarcodeSiteVariant(UBound(BarcodeStrInfor) - i) = BarcodeStrInfor(i)
                    '20240502 Mason - Store Barcode to array by site
                    IEDAInputStrArr(UBound(BarcodeStrInfor) - i) = BarcodeStrInfor(i)
                End If
            Next i
            
            '20240502 Mason - String processing
            For i = 0 To UBound(IEDAInputStrArr)
                If i = UBound(IEDAInputStrArr) Then
                    IEDAInputStr = IEDAInputStr & IEDAInputStrArr(i)
                Else
                    IEDAInputStr = IEDAInputStr & IEDAInputStrArr(i) & ","
                End If
            Next i
            '20240502 Mason - Save value to registry
            Call IEDA_SaveRegistry(IEDAInputStr, "Handler2DID")
            
            '20240502 Mason - Print to text file----------------
'            Dim Current_Dirctory As String
'            Dim Folder_Name As String
'
'            Folder_Name = "BarCodeData" & "_" & right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2) & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2)
'            Current_Dirctory = CurDir & "\" & Folder_Name
'
'            Dim TempStr As String
'            TempStr = Current_Dirctory & "\"
'            Dim fso As New FileSystemObject
'            If Dir(TempStr, vbDirectory) = Empty Then
'                MkDir TempStr
'            End If
'            fso.CreateTextFile TempStr & "Barcode.txt", True
'            Dim s As String
'            Dim n As Integer
'            n = FreeFile()
'            Open TempStr & "Barcode.txt" For Append As n
'            Print #n, IEDAInputStr
'            Close #n
            '20240502 Mason - Print to text file----------------
            
            TheExec.Datalog.WriteComment "Handler2DID: " & IEDAInputStr
            
            '20231219 Mason - Add test limit to print the barcode to stdf
            For Each site In TheExec.Sites
                TheExec.flow.TestLimit 0, Tname:=BarcodeSiteVariant
            Next
            
            TheExec.Datalog.WriteComment "---------Capture 2DID Information Complete -------------"
            
            BarcodeString = ""
        Else
            TheExec.Datalog.WriteComment "Please check the IC has the 2DID information"
        End If
    Else
        TheExec.Datalog.WriteComment "IDString = Empty"
    End If
    
    
'    end_time = TheExec.Timer
'    test_time = end_time - start_time
'    TheExec.Datalog.WriteComment "Test time=" & test_time
    
 End Function

Public Function EEPROM_Write_Hw(Optional str_data As String)
    
    Dim REC() As IDIB_EEPROM_RecordObj
    Dim col_readback As IDIB_EEPROM_RecordObj
    Dim config(2) As Long, i As Integer
    Dim ReadBack_String As String
    'Dim i As Integer
    Call EEPROM_init
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "================================================="
    TheExec.Datalog.WriteComment "WRITE DIB INFO TO EEPROM "
    TheExec.Datalog.WriteComment "================================================="
    If UCase(TheHdw.Tester.Type) = UCase("Jaguar") Then
        With TheHdw.DIB.EEPROM
            TheHdw.DIB.EEPROM.Record.ReadFromHW
            REC = TheHdw.DIB.EEPROM.Record.List
            
            If LBound(REC) = 0 And UBound(REC) = -1 Then
                TheExec.Datalog.WriteComment "ID = " & glb_EEPROM_recordID
                TheExec.Datalog.WriteComment "Value = " & "cold temp"
                .Record(glb_EEPROM_recordID) = "cold temp"
                .Record.WriteToHW
            Else
                
                For i = LBound(REC) To UBound(REC)
                    TheHdw.DIB.EEPROM.Record.Delete REC(i).ID, REC
                Next i
                
            End If
        End With

        
    End If
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "================================================="
    TheExec.Datalog.WriteComment "READ BACK DIB INFO FROM EEPROM "
    TheExec.Datalog.WriteComment "================================================="
    If UCase(TheHdw.Tester.Type) = UCase("Jaguar") Then
        TheHdw.DIB.EEPROM.Record.ReadFromHW
        REC = TheHdw.DIB.EEPROM.Record.List
        
        If LBound(REC) = 0 And UBound(REC) = -1 Then
            TheExec.Datalog.WriteComment "EEPROM.Record.List is empty!!"
            TheExec.Datalog.WriteComment "ERROR!! Write EEPROM Failed!"
            TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
            Exit Function
        Else
            With TheHdw.DIB.EEPROM
                For i = LBound(REC) To UBound(REC)
                    If REC(i).ID = glb_EEPROM_recordID Then
                       
                        If .Record(glb_EEPROM_recordID) = "cold temp" Then
                            TheExec.Datalog.WriteComment "EEPROM: " & .Record(glb_EEPROM_recordID)
                        Else
                            TheExec.Datalog.WriteComment "ERROR: EEPROM:" & .Record(glb_EEPROM_recordID)
                            TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
                            
                        End If
                    Else
                        TheExec.Datalog.WriteComment "EEPROM: " & glb_EEPROM_recordID
                    End If
                Next i
            End With
        End If
    ElseIf UCase(TheHdw.Tester.Type) = UCase("UltraFLEXplus") Then
        TheHdw.DIB.EEPROM.Record.ReadFromHW
        REC = TheHdw.DIB.EEPROM.Record.List
        
        If LBound(REC) = 0 And UBound(REC) = -1 Then
            TheExec.Datalog.WriteComment "EEPROM.Record.List is empty!!"
            TheExec.Datalog.WriteComment "ERROR!! Write EEPROM Failed!"
            TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
            Exit Function
        Else
            With TheHdw.DIB.Segments.Item(1).EEPROM
                For i = LBound(REC) To UBound(REC)
                    If REC(i).ID = glb_EEPROM_recordID Then
                       
                        If .Record(glb_EEPROM_recordID) = "cold temp" Then
                            TheExec.Datalog.WriteComment "EEPROM: " & .Record(glb_EEPROM_recordID)
                        Else
                            TheExec.Datalog.WriteComment "ERROR: EEPROM:" & .Record(glb_EEPROM_recordID)
                            TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
                            
                        End If
                    Else
                        TheExec.Datalog.WriteComment "EEPROM: " & glb_EEPROM_recordID
                    End If
                Next i
            End With
            
        End If
    
    End If
    
End Function


Public Function EEPROM_init()

    Dim config(2) As Long
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "================================================="
    TheExec.Datalog.WriteComment "EEPROM Init "
    TheExec.Datalog.WriteComment "================================================="
    config(0) = 32768
    config(1) = 0
    config(2) = 0
    If UCase(TheHdw.Tester.Type) = UCase("Jaguar") Then

        If TheHdw.DIB.EEPROM.IsProgrammed = False Then
            TheHdw.DIB.EEPROM.Program (config)
        End If
        TheExec.Datalog.WriteComment "Programs and Formats DIB EEPROMs Successful!!"
'    ElseIf UCase(thehdw.Tester.type) = UCase("UltraFLEXplus") Then
'         thehdw.DIB.Segments("1").EEPROM.program
    Else
        TheExec.Datalog.WriteComment "Tester Type NOT UltraFLEX"
        Exit Function
    End If
    
End Function

Public Function Read_EEPROM_proberTemp_compare() As Long
    
    Dim CommandBuf As String
    Dim ReplyBuf As String
    Dim StartTime As Double
    Dim endTime As Double, dl_temp As Double
    Dim eeprom_read_data As String
    Dim bl_GPIB_Connect As Boolean
    Dim site As Variant
    
    If GPIBCtrl.GPIB_Connect Then
    ''do nothing
    Else
        TheExec.Datalog.WriteComment "Fail!!! GPIB Connect issue please check register edit path!!"
        TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
        Exit Function
    End If
    
    If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_CP_*" Then
        CommandBuf = "f1"
    Else
        ''FT do nothing
        'CommandBuf = "TEMPARM?"
        TheExec.Datalog.WriteComment "CurrentJob is not or WLFT!!!"
        TheExec.Datalog.WriteComment "By pass temperture check!!!"
        Exit Function
    End If
    
    Call GPIBSendQuery_NI(CommandBuf, ReplyBuf)
    dl_temp = CVar(ReplyBuf)
    
    ''EEPROM
    eeprom_read_data = EEPROM_readBack()
    
    If eeprom_read_data = "cold temp" Then
   
        If dl_temp < 35 Then
        ''pass
            TheExec.Datalog.WriteComment "Pass eeprom_read_data: " & eeprom_read_data & " temp: " & dl_temp
            TheExec.flow.TestLimit resultVal:=-1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
            'Debug.Print "Pass eeprom_read_data: " & eeprom_read_data & " dl_temp: " & dl_temp
        Else
        ''fail
            TheExec.Datalog.WriteComment "Fail eeprom_read_data: " & eeprom_read_data & " temp: " & dl_temp
            TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
            'Debug.Print "Fail eeprom_read_data: " & eeprom_read_data & " dl_temp: " & dl_temp
            'MsgBox ("Fail!!! eeprom read back = cold temp!!! Prober: TEMP = " & dl_temp)
            'Stop
        End If
    
    ElseIf eeprom_read_data = "hot temp" Then
        If dl_temp > 15 Then
        ''pass
            TheExec.Datalog.WriteComment "Pass eeprom_read_data: " & eeprom_read_data & " temp: " & dl_temp
            TheExec.flow.TestLimit resultVal:=-1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
            'Debug.Print "Pass eeprom_read_data: " & eeprom_read_data & " dl_temp: " & dl_temp
        Else
        ''fail
            TheExec.Datalog.WriteComment "Fail eeprom_read_data: " & eeprom_read_data & " temp: " & dl_temp
            TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
            'Debug.Print "Fail eprom_read_data: " & eeprom_read_data & " dl_temp: " & dl_temp
            'MsgBox ("Fail!!! eeprom read back = hot temp!!! Prober: TEMP = " & dl_temp)
            'Stop
        End If
    Else
    ''do nothing
    End If
    
End Function

Public Function EEPROM_readBack() As String
    Dim REC() As IDIB_EEPROM_RecordObj
    Dim col_readback As IDIB_EEPROM_RecordObj
    Dim i As Integer
    Dim ReadBack_String As String
    
    Call EEPROM_init
    EEPROM_readBack = EEPROM_read()
    
End Function

Public Function EEPROM_read()

    Dim REC() As IDIB_EEPROM_RecordObj
    Dim col_readback As IDIB_EEPROM_RecordObj
    Dim i As Integer, EEPROM_Read_temp As String
    Dim ReadBack_String As String
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "================================================="
    TheExec.Datalog.WriteComment "READ BACK DIB INFO FROM EEPROM "
    TheExec.Datalog.WriteComment "================================================="
    If UCase(TheHdw.Tester.Type) = UCase("Jaguar") Then
        TheHdw.DIB.EEPROM.Record.ReadFromHW
        REC = TheHdw.DIB.EEPROM.Record.List
        
        If LBound(REC) = 0 And UBound(REC) = -1 Then
            ''check EEPROM working
            TheHdw.DIB.EEPROM.Record("test") = "work"
            TheHdw.DIB.EEPROM.Record.WriteToHW
            TheHdw.DIB.EEPROM.Record.ReadFromHW
            REC = TheHdw.DIB.EEPROM.Record.List
            
            If LBound(REC) = 0 And UBound(REC) = -1 Then
                ''EEPROM broken
                TheExec.Datalog.WriteComment "ERROR!! Not working!"
                TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
                Exit Function
            Else
                ''EEPROM is hot Temp,
                REC = TheHdw.DIB.EEPROM.Record.List
                TheHdw.DIB.EEPROM.Record.Delete "test", REC
                TheHdw.DIB.EEPROM.Record.WriteToHW
                
                TheExec.Datalog.WriteComment "EEPROM didn't have temp information!!"
                EEPROM_read = "hot temp"
                TheExec.Datalog.WriteComment "EEPROM: hot temp"
                Exit Function
            
            End If
        Else
            With TheHdw.DIB.EEPROM
                For i = LBound(REC) To UBound(REC)
                    If REC(i).ID = glb_EEPROM_recordID Then
                       
                        If .Record(glb_EEPROM_recordID) = "cold temp" Then
                            TheExec.Datalog.WriteComment "EEPROM: " & .Record(glb_EEPROM_recordID)
                            EEPROM_read = "cold temp"
                        Else
                            TheExec.Datalog.WriteComment "ERROR: EEPROM:" & .Record(glb_EEPROM_recordID)
                            TheExec.flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, ForceResults:=tlForceNone
                            
                        End If
                    Else
                        TheExec.Datalog.WriteComment "EEPROM: " & glb_EEPROM_recordID
                        EEPROM_read = "hot temp"
                    End If
                Next i
            End With
        End If
    End If
End Function

Public Function GetStoredData(keyname As String) As Variant
    keyname = LCase(keyname)
    If Not gDictSiteLong.Exists(keyname) Then
        TheExec.ErrorLogMessage "Stored capture data " & keyname & " not found."
    Else
        Set GetStoredData = gDictSiteLong(keyname)
    End If
End Function
