VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HRAMInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private site_ As Integer
Private voltage_ As String
Private pattern_ As String
Private failPins_() As String
Private failCount_ As Long
Private LHReceived_ As Collection
Private vectors_ As Collection
Private cycles_ As Collection
Private scans_ As Collection
Private pins_ As Collection


Sub Class_Initialize()
On Error GoTo errHandler
    site_ = 0
    voltage_ = vbNullString
    pattern_ = vbNullString
    failCount_ = 0
    Set LHReceived_ = New Collection
    Set vectors_ = New Collection
    Set cycles_ = New Collection
    Set scans_ = New Collection
    Set pins_ = New Collection
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "HRAMInfo", "Class_Initialize")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

''' ===== write only for below ===== '''
''' =============================== '''

''' ===== read only for below ===== '''
Public Property Get site() As Integer
    site = site_
End Property
Public Property Get Voltage() As String
    Voltage = voltage_
End Property
Public Property Get Pattern() As String
    Pattern = pattern_
End Property
Public Property Get FailCount() As Long
    FailCount = failCount_
End Property
Public Property Get LHReceived() As Collection
    Set LHReceived = LHReceived_
End Property
Public Property Get vectors() As Collection
    Set vectors = vectors_
End Property
Public Property Get cycles() As Collection
    Set cycles = cycles_
End Property
Public Property Get Scans() As Collection
    Set Scans = scans_
End Property
Public Property Get Pins() As Collection
    Set Pins = pins_
End Property
''' =============================== '''

Public Function patGenInfoUF(failNum As Long, ByRef pgCyclesSV As SiteVariant, ByRef pgVecorSV As SiteVariant, ByRef pgScanIdxsSV As SiteVariant)
On Error GoTo errHandler
    Dim num As Long
        Dim cycles() As Double
    Dim vectors() As Long
    Dim ScanIdxs() As Long
    ReDim cycles(failNum - 1)
    ReDim vectors(failNum - 1)
    ReDim ScanIdxs(failNum - 1)
    
    For num = 0 To failNum - 1
        cycles(num) = TheHdw.Digital.hram.PatGenInfo(num, pgCycle)
        vectors(num) = TheHdw.Digital.hram.PatGenInfo(num, pgVector)
        ScanIdxs(num) = TheHdw.Digital.hram.PatGenInfo(num, pgScanIndex)
    Next num
    
    If failNum = 1 Then
        pgCyclesSV = cycles(0)
        pgVecorSV = vectors(0)
        pgScanIdxsSV = ScanIdxs(0)
    Else
        pgCyclesSV = cycles
        pgVecorSV = vectors
        pgScanIdxsSV = ScanIdxs
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "HRAMInfo", "patGenInfoUF")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function captureFailInfo(siteNumber As Integer, Pat As String, vol As String) As Boolean
    ''' capture fail info from HRAM '''
On Error GoTo errHandler

    'TheHdw.StartStopwatch
    Dim num As Long
    Dim failNum As Long
    Dim pgCyclesSV As New SiteVariant
    Dim pgVecorSV As New SiteVariant
    Dim pgScanIdxsSV As New SiteVariant
    Dim PinPF As New PinListData
    Dim PinData As New PinListData
    Dim Pins As String
    Dim LHs As String
    Dim IGXLversion As String
    Dim versionArray() As String
    Dim pin As Variant
    
    site_ = siteNumber
    voltage_ = vol
    pattern_ = Pat
    failCount_ = 0
    
    'sleep (1000) 'DBDP crash issue caused by hram getCycle timeout
    'check IGXL version
    'run sleep(1000) if IGXL version is older then 10.30(10.30.90)
'        IGXLversion = TheExec.SoftwareVersion
'    versionArray = Split(IGXLversion, ".")
'    If (CInt(versionArray(0)) < 10) Then
'        sleep (1000)
'    ElseIf (CInt(versionArray(0)) = 10) And (versionArray(1) < 30) Then
'        sleep (1000)
'    End If

    If glb_TesterType = "UltraFLEXplus" Then sleep (1000)
        
    failNum = TheHdw.Digital.hram.CapturedCycles
    failPins_ = TheHdw.Digital.FailedPins(siteNumber) ''' get fail pin

    If glb_TesterType = "Jaguar" Then  ' is uf
        patGenInfoUF failNum, pgCyclesSV, pgVecorSV, pgScanIdxsSV
    Else
        pgCyclesSV = TheHdw.Digital.hram.PatGenInfo(0, pgCycle, failNum)
        pgVecorSV = TheHdw.Digital.hram.PatGenInfo(0, pgVector, failNum)
        pgScanIdxsSV = TheHdw.Digital.hram.PatGenInfo(0, pgScanIndex, failNum)
    End If
    
    If failNum = 0 Or (Not Not failPins_) = 0 Then
        captureFailInfo = False: Exit Function
    End If
    
    PinPF = TheHdw.Digital.Pins(Join(failPins_, ",")).hram.PinPF(0, 0, failNum)
    PinData = TheHdw.Digital.Pins(Join(failPins_, ",")).hram.PinData(0, 0, failNum) '2nd '0' for NON-2x/4x
    
    For num = 0 To failNum - 1
        Pins = vbNullString 'value must be reset
        LHs = vbNullString
        
        For Each pin In failPins_
            If failNum = 1 Then
                If PinPF.Pins(pin)(siteNumber) = tlResultFail Then
                    Pins = Pins + CStr(pin) + "/"
                    LHs = LHs + CStr(PinData.Pins(pin)(siteNumber)) + "/"
                End If
            Else
                If PinPF.Pins(pin)(siteNumber)(num) = tlResultFail Then
                    Pins = Pins + CStr(pin) + "/"
                    LHs = LHs + CStr(PinData.Pins(pin)(siteNumber)(num)) + "/"
                End If
            End If
        Next pin
        
        If Pins <> "" Then
            pins_.Add left(Pins, Len(Pins) - 1), str(failCount_)
            LHReceived_.Add left(LHs, Len(LHs) - 1), str(failCount_)
            
            If failNum = 1 Then
                cycles_.Add CStr(pgCyclesSV(siteNumber)), str(failCount_) ' it will be a scalar value, not array, so cannot index
                vectors_.Add CStr(pgVecorSV(siteNumber)), str(failCount_)
                scans_.Add CStr(pgScanIdxsSV(siteNumber)), str(failCount_)
            Else
                cycles_.Add CStr(pgCyclesSV(siteNumber)(num)), str(failCount_)
                vectors_.Add CStr(pgVecorSV(siteNumber)(num)), str(failCount_)
                scans_.Add CStr(pgScanIdxsSV(siteNumber)(num)), str(failCount_)
            End If
            
            failCount_ = failCount_ + 1
        End If
    Next num
    
    captureFailInfo = True
          'TheExec.AddOutput TheHdw.ReadStopwatch
    Exit Function

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "HRAMInfo", "captureFailInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function
