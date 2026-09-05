VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LVCC_VminBoundary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum Headers
    hName = 0
    hFailCount = 1
    hSite = 2
    hVoltage = 3
    hTestSuite = 4
    hPattern = 5
    hscan = 6
    hRelative = 7
    hCycle = 8
    hPins = 9
    hLH = 10
End Enum
Private Const HEADER_DFC = ",<failcycle#>,<site>,<voltage>,<testsuite>,<pattern>,<scan>,<relative>,<cycle>,<pins>,<LH received>" ''' header for first printing
Private Const ENABLE_DFC = "Enable_DFTLHFC_Debug"
Private Const ENABLE_FAILLOG = "Enable_Faillog_Debug"
Private Const HEADERNAME_DFC = "DFC"
Private Const HEADER_COUNT = 10 ''' sum of enum headers
'Chris modify 220418
Private LVCC_VOLTAGE As Double '= 0.01 ''' the voltage step condition for faillog
Private LVCC_COUNT As Double '= 2 ''' capture only one count, must be capture count + 1
Private LVCC_Value As Double
Private FAIL_COND As Double
'''''''''''''''
Private Const ONE_POINT = False
Private isStoreAnyHRAM_ As Boolean
Private headerPrint_ As Boolean
Private instanceName_ As String
Private instanceHead_ As String
Private testSuite_ As String
Private hramDatas_ As Dictionary

Sub Class_Initialize()
On Error GoTo errHandler
    '' set capture type: fail, capture cycle: 512, trigger: fail
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    If TheExec.TesterMode = testModeOffline Then TheHdw.Digital.hram.size = 512 * 2
    If TheExec.TesterMode = testModeOnline Then TheHdw.Digital.hram.size = gl_HRAMmaxDepth
    'If TheExec.TesterMode = testModeOnline Then TheHdw.Digital.hram.size = TheHdw.Digital.hram.maxDepth
    TheHdw.Digital.hram.CaptureType = captFail
    TheHdw.Digital.hram.SetTrigger trigFail, False, 0, True
    TheHdw.Digital.Patgen.ClearFail
    '' setup for ffc info
    instanceName_ = UCase(TheExec.DataManager.instancename)
    instanceHead_ = HEADERNAME_DFC
    testSuite_ = getTestSuitName
    isStoreAnyHRAM_ = False
    headerPrint_ = False
    Set hramDatas_ = New Dictionary
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary", "Class_Initialize")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
''' ===== write only for below ===== '''
''' =============================== '''

''' ===== read only for below ===== '''
Public Property Get IsAnySiteFail() As Boolean
    IsAnySiteFail = TheHdw.Digital.Patgen.PatternBurstPassed
End Property
Public Property Get IsEnableDFTLHFC() As Boolean
    IsEnableDFTLHFC = TheExec.enableWord(ENABLE_DFC)
End Property
Public Property Get IsEnableFAILLOG() As Boolean
    IsEnableFAILLOG = TheExec.enableWord(ENABLE_FAILLOG)
End Property
Public Property Get IsStoreAnyHRAM() As Boolean
    IsStoreAnyHRAM = isStoreAnyHRAM_
End Property
Public Property Get HRAMDatas() As Dictionary
    Set HRAMDatas = hramDatas_
End Property
Public Property Let LvccValue(m_LVCC_VAL As Double)
    LVCC_Value = m_LVCC_VAL
End Property
Public Property Let LvccVoltage(m_LVCC_VOLTAGE As Double)
    LVCC_VOLTAGE = m_LVCC_VOLTAGE
End Property
Public Property Get LvccVoltage() As Double
    LvccVoltage = LVCC_VOLTAGE
End Property
Public Property Let LvccCount(m_LVCC_COUNT As Integer)
    LVCC_COUNT = m_LVCC_COUNT
End Property
Public Property Get LvccCount() As Integer
    LvccCount = LVCC_COUNT
End Property
Public Property Get IsOnlyOnePoint() As Boolean
    IsOnlyOnePoint = ONE_POINT
End Property
''' =============================== '''

Private Function getTestSuitName() As String
On Error GoTo errHandler
    ''' initial DFC name for each variable '''
    If instanceName_ = "" Then instanceHead_ = "NotFound!": Exit Function
    Dim instinfo() As String
    Dim clk_info_temp As String                                     '220315 edit
    Dim i As Integer
    instinfo = Split(instanceName_, "_")
    instanceHead_ = HEADERNAME_DFC
    ReDim Preserve instinfo(9)                                      '220315 edit
    clk_info_temp = instinfo(8)                                     '220316 edit
    If UCase(clk_info_temp) Like "PADIO*" Then
        For i = 1 To 9: instinfo(i - 1) = instinfo(i): Next i       '220315 edit
        instinfo(3) = clk_info_temp                                 '220315 edit
        ReDim Preserve instinfo(3)                                  '220315 edit
    ElseIf UCase(clk_info_temp) Like "XIO*" Then
        For i = 1 To 9: instinfo(i - 1) = instinfo(i): Next i       '220315 edit
        instinfo(3) = clk_info_temp                                 '220315 edit
        ReDim Preserve instinfo(3)
    ElseIf UCase(clk_info_temp) Like "TCK*" Then
        For i = 1 To 9: instinfo(i - 1) = instinfo(i): Next i       '220315 edit
        instinfo(3) = clk_info_temp                                 '220315 edit
        ReDim Preserve instinfo(3)
    ElseIf UCase(clk_info_temp) Like "SHIFTIN*" Then
        For i = 1 To 9: instinfo(i - 1) = instinfo(i): Next i       '220315 edit
        instinfo(3) = clk_info_temp                                 '220315 edit
        ReDim Preserve instinfo(3)
    ElseIf UCase(clk_info_temp) Like "RT*" Then
        For i = 1 To 9: instinfo(i - 1) = instinfo(i): Next i       '220315 edit
        instinfo(3) = clk_info_temp                                 '220315 edit
        ReDim Preserve instinfo(3)
    ElseIf UCase(clk_info_temp) Like "GPIO*" Then
        For i = 1 To 9: instinfo(i - 1) = instinfo(i): Next i       '220315 edit
        instinfo(3) = clk_info_temp                                 '220315 edit
        ReDim Preserve instinfo(3)
    ElseIf UCase(clk_info_temp) Like "CLK*" Then
        For i = 1 To 9: instinfo(i - 1) = instinfo(i): Next i       '220331 edit
        instinfo(3) = clk_info_temp                                 '220331 edit
        ReDim Preserve instinfo(3)
    Else                                                            '220316 edit
        ReDim Preserve instinfo(3)
        For i = 1 To 3: instinfo(i - 1) = instinfo(i): Next i
        ReDim Preserve instinfo(2)
    End If                                                          '220316 edit
    getTestSuitName = Join(instinfo, "_")
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary", "getTestSuitName")
    If AbortTest Then Exit Function Else Resume Next
End Function
Private Function SliceArray(ary() As String, head As Integer, rail As Integer) As Variant
On Error GoTo errHandler
    Dim out As Variant
    Dim idx As Integer
    out = ary
    If head = 0 Then
        ReDim Preserve out(rail)
    Else
        For idx = head To rail
            out(idx - head) = out(idx)
        Next idx
        ReDim Preserve out(rail - head)
    End If
    SliceArray = out
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary", "SliceArray")
    If AbortTest Then Exit Function Else Resume Next
End Function
Private Function SplitSentence(info() As String) As Variant
On Error GoTo errHandler
    Dim tmp As Variant
    Dim out As Variant
    Dim outs() As String
    Dim LHs() As String
    Dim pins() As String
    Dim idx As Integer
    Dim mid As Integer
    Dim pivot As Integer
    Dim Count As Integer
    Dim left As Integer
    Dim right As Integer
    Dim index As Integer
    Dim dyn_len As Integer
    Dim fix_len As Integer
    
    tmp = SliceArray(info, 0, HEADER_COUNT - 2)
    fix_len = Len(Join(tmp, ","))
    LHs = Split(info(HEADER_COUNT), "/")
    pins = Split(info(HEADER_COUNT - 1), "/")

    pivot = 0: index = 0: Count = UBound(pins)
    ''' split the combination of pins and corresponding outputs which length > 255
    While pivot <= Count
        left = pivot
        right = Count
        ''' find the index that let length close to 255
        ''' note: it may be higher or lower than 255 "-----(lower)---(255)---(higher)-----"
        While left <= right
            mid = Int((left + right) / 2)
            dyn_len = Len(Join(SliceArray(LHs, pivot, mid), "/")) + Len(Join(SliceArray(pins, pivot, mid), "/"))
            If fix_len + dyn_len + 2 > 255 Then ''' 2 is for comma
                right = mid - 1
            Else
                left = mid + 1
            End If
        Wend
        ''' re-calculate the case that higher than 255 words, reduce its index value to match < 255
        Do While mid > pivot
            If Not fix_len + dyn_len + 2 > 255 Then Exit Do
            mid = mid - 1
            dyn_len = Len(Join(SliceArray(LHs, pivot, mid), "/")) + Len(Join(SliceArray(pins, pivot, mid), "/"))
        Loop
        ''' store the new sentence which length <= 255
        ReDim Preserve outs(index)
        outs(index) = Join(tmp, ",") + "," + Join(SliceArray(pins, pivot, mid), "/") + "," + Join(SliceArray(LHs, pivot, mid), "/")
        pivot = mid + 1: index = index + 1
    Wend
    SplitSentence = outs
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary", "SplitSentence")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function initialFFC()
On Error GoTo errHandler
    ''' initial the hram '''
    Set hramDatas_ = Nothing
    Set hramDatas_ = New Dictionary
    isStoreAnyHRAM_ = False
    headerPrint_ = False
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary", "initialFFC")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function storeHRAMData(pat As String, shmooPins() As String, site As Variant) As Boolean
On Error GoTo errHandler
    ''' capture hram and store into data structure '''
    Dim HRAMData As HRAMInfo
'    Dim site As Variant
    Dim Voltage As String
'    For Each site In TheExec.sites.Active
        If TheHdw.Digital.Patgen.PatternBurstPassedPerSite(site) Then GoTo nextSite
        Set HRAMData = New HRAMInfo
        FAIL_COND = TheHdw.DCVS.pins(shmooPins(0)).Voltage.Alt.value
        Voltage = Join(shmooPins, "/") + ":" + Format(FAIL_COND, "0.000")
        If Not HRAMData.captureFailInfo(Int(site), pat, Voltage) Then
            isStoreAnyHRAM_ = False
            Set HRAMData = Nothing
            storeHRAMData = False
            Exit Function
        End If
        ''' if not capture from new site, create one for it
        If Not hramDatas_.Exists(site) Then
            Dim info As Collection
            Set info = New Collection
            info.Add HRAMData
            hramDatas_.Add site, info
            Set info = Nothing
        Else
            hramDatas_(site).Add HRAMData
        End If
        Set HRAMData = Nothing
nextSite:
'    Next site
    TheHdw.Digital.Patgen.ClearFail ''' reset HRAM
    If hramDatas_.Count > 0 Then isStoreAnyHRAM_ = True
    storeHRAMData = True
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary", "storeHRAMData")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function processFailCycles() As Boolean
On Error GoTo errHandler
    ''' not used for now '''
    ''' capture fail cycle and transfer to FFC format '''
    Dim site As Variant
    For Each site In TheExec.sites.Active
        If Not hramDatas_.Exists(site) Then GoTo nextSite
        ''' TODO
        ''' print out the current voltage for fail log
        printFailInfo (Int(site))
nextSite:
    Next site
    processFailCycles = True
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary", "processFailCycles")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function printFailInfo()
On Error GoTo errHandler
    ''' print the FFC info on datalog '''
    Dim num As Integer
    Dim out As Variant
    Dim data As Variant
    Dim hram As Variant
    Dim Output As String
    Dim outs() As String
    Dim info(HEADER_COUNT) As String
    
    If do_shmoo_hole_flag = True Then
        instanceHead_ = "Shmoo hole" '20231018
    Else
        instanceHead_ = "DFC"
    End If
    
    If Not headerPrint_ Then
        TheExec.Datalog.WriteComment instanceHead_ + HEADER_DFC
        headerPrint_ = True
    End If
    
    For Each hram In hramDatas_.Items
        For Each data In hram
            info(Headers.hName) = Trim(instanceHead_)
            'info(Headers.hTestSuite) = Trim(testSuite_) & ":" & Format(FAIL_COND - LVCC_Value, "0.000") 'Chris added 220418
            info(Headers.hTestSuite) = Trim(TheExec.DataManager.instancename) & ":" & Format(FAIL_COND - LVCC_Value, "0.000") 'Charles changed format 231025
            info(Headers.hSite) = Trim(data.site)
            info(Headers.hVoltage) = Trim(data.Voltage)
            info(Headers.hPattern) = Trim(data.Pattern)
            
            For num = 0 To data.FailCount - 1
                info(Headers.hFailCount) = Trim(str(num + 1))
                info(Headers.hscan) = Trim(str(data.Scans(str(num))))
                info(Headers.hRelative) = Trim(str(data.vectors(str(num))))
                info(Headers.hCycle) = Trim(str(data.cycles(str(num))))
                info(Headers.hPins) = Trim(data.pins(str(num)))
                info(Headers.hLH) = Trim(data.LHReceived(str(num)))
                Output = Join(info, ",")
                If Len(Output) > 255 Then
                    outs = SplitSentence(info)
                    For Each out In outs
                        TheExec.Datalog.WriteComment out
                    Next out
                Else
                    TheExec.Datalog.WriteComment Output
                End If
            Next num
        Next data
    Next hram
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary", "printFailInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function



