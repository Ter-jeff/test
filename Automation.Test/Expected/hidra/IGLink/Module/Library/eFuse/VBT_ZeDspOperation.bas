Attribute VB_Name = "VBT_ZeDspOperation"
Option Explicit
Private Const RunAtDsp = True
Private site As Variant

Public Sub RunDspSet()
On Error GoTo errHandler
    ''202005xx for ap
    If TheHdw.DSP.ExecutionMode <> tlDSPModeAutomatic Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic 'tlDSPModeForceAutomatic
    End If
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "RunDspSet")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_SingleDoubleCal(ByVal inWave1 As DSPWave, ByVal inWave2 As DSPWave, ByRef result As SiteLong)
On Error GoTo errHandler
Dim comp As New DSPWave

    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.SingleDoubleCal inWave1, inWave2, result
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    For Each site In TheExec.sites
        Set comp = inWave1.LogicalCompare(EqualTo, inWave2)
        result = comp.CountElements(EqualTo, 0)
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_SingleDoubleCal")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Public Sub Ze_DspwaveToSerialMsb(ByRef InWave As DSPWave, ByVal parallelBits As Long)
'If Not RunAtDsp Then GoTo s_SERIAL
'     'Call RunDspSet
'     rundsp.DspwaveToSerialMsb InWave, parallelBits
'     'Call RunDspReSet
'Exit Sub
's_SERIAL:
'        For Each site In TheExec.sites
'                InWave = InWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, parallelBits, 0, Bit0IsMsb)
'        Next
'End Sub

'Public Sub Ze_DspwaveToSerialLsb(ByRef InWave As DSPWave, ByVal parallelBits As Long)
'If Not RunAtDsp Then GoTo s_SERIAL
'     'Call RunDspSet
'     rundsp.DspwaveToSerialLsb InWave, parallelBits
'     'Call RunDspReSet
'Exit Sub
's_SERIAL:
'        For Each site In TheExec.sites
'                InWave = InWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, parallelBits, 0, Bit0IsLsb)
'        Next
'End Sub

Public Sub Ze_DspwaveToParallelMsb(ByVal inwave As DSPWave, ByVal parallelBits As Long, ByRef outwave As DSPWave)
On Error GoTo errHandler
Dim appSize As Long
Dim appWave As New DSPWave

    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.DspwaveToParallelMsb inwave, parallelBits, outwave
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    appSize = GlbUtility.GetSampleSize(inwave) Mod parallelBits
    If appSize <> 0 Then
        appSize = parallelBits - appSize
        appWave.CreateConstant 0, appSize, inwave.DataType
    End If
    For Each site In TheExec.sites
        If appSize <> 0 Then inwave = inwave.Concatenate(appWave)
        outwave = inwave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, parallelBits, 0, Bit0IsMsb)
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_DspwaveToParallelMsb")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Public Sub Ze_DspwaveToParallelLsb(ByVal InWave As DSPWave, ByVal parallelBits As Long, ByRef outwave As DSPWave)
'If Not RunAtDsp Then GoTo s_SERIAL
'     'Call RunDspSet
'     rundsp.DspwaveToParallelLsb InWave, parallelBits, outwave
'     'Call RunDspReSet
'Exit Sub
's_SERIAL:
'    Dim appSize As Long
'     appSize = GlbUtility.GetSampleSize(InWave) Mod parallelBits
'     If appSize <> 0 Then
'            appSize = parallelBits - appSize
'            Dim appWave As New DSPWave
'            appWave.CreateConstant 0, appSize, InWave.DataType
'     End If
'
'        For Each site In TheExec.sites
'                If appSize <> 0 Then InWave = InWave.Concatenate(appWave)
'                outwave = InWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, parallelBits, 0, Bit0IsLsb)
'        Next
'End Sub

Public Sub Ze_XExpandBitVector(ByVal inwave As DSPWave, ByVal dealVecRpt As Long, ByRef outwave As DSPWave)
On Error GoTo errHandler

    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.XExpandBitVector inwave, dealVecRpt, outwave
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    For Each site In TheExec.sites
        outwave = inwave.repeat(dealVecRpt).MatrixTranspose(dealVecRpt)
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_XExpandBitVector")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_DspwaveParallel8ToParallel32(ByVal inwave As DSPWave, ByRef outwave As DSPWave)
On Error GoTo errHandler
Dim i As Long, SampleSize As Long

    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.DspwaveParallel8ToParallel32 inwave, outwave
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    outwave.CreateConstant 0, SampleSize / 4
    For Each site In TheExec.sites
        For i = 0 To SampleSize \ 4 - 1
            outwave.ElementLite(i) = inwave.ElementLite(4 * i + 3) * 2 ^ 24 + inwave.ElementLite(4 * i + 2) * 2 ^ 16 + inwave.ElementLite(4 * i + 1) * 2 ^ 8 + inwave.ElementLite(4 * i)
        Next i
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_DspwaveParallel8ToParallel32")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_DspwaveBlankCheck(ByRef inwave As DSPWave, ByRef FBC As SiteLong)
On Error GoTo errHandler

    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.DspwaveBlankCheck inwave, FBC
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    For Each site In TheExec.sites
        FBC = inwave.CountElements(GreaterThan, 0)
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_DspwaveBlankCheck")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_SplitWaveSerial32(ByVal inwave As DSPWave, ByRef out1 As DSPWave, ByRef out2 As DSPWave)
On Error GoTo errHandler
Dim i As Long, j As Long, SampleSize As Long 'assume 32 bits always

    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.SplitWaveSerial32 inwave, out1, out2
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    out1.CreateConstant 0, SampleSize / 2
    out2.CreateConstant 0, SampleSize / 2
    For Each site In TheExec.sites
        For i = 0 To (SampleSize \ 32 - 1)
            For j = 0 To 15
                out1.ElementLite(16 * i + j) = inwave.ElementLite(32 * i + j)
                out2.ElementLite(16 * i + j) = inwave.ElementLite(32 * i + j + 16)
            Next
        Next i
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_SplitWaveSerial32")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_SplitWave32(ByVal inwave As DSPWave, ByRef out1 As DSPWave, ByRef out2 As DSPWave)
On Error GoTo errHandler
Dim i As Long, SampleSize As Long, twos As Double, value As Double   'assume 32 bits always
    
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.SplitWave32 inwave, out1, out2
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    out1.CreateConstant 0, SampleSize
    out2.CreateConstant 0, SampleSize
    twos = 2# ^ 32
    For Each site In TheExec.sites
        For i = 0 To (SampleSize - 1)
        value = inwave.ElementLite(i)
        If value < 0 Then value = value + twos
            out1.ElementLite(i) = Fix(value / 2# ^ 16)
            out2.ElementLite(i) = value - out1.ElementLite(i) * (2# ^ 16)
        Next i
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_SplitWave32")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
 
Public Sub Ze_SplitWave8(ByVal inwave As DSPWave, ByRef out1 As DSPWave, ByRef out2 As DSPWave)
On Error GoTo errHandler
Dim i As Long, SampleSize As Long  'assume 32 bits always
    
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.SplitWave8 inwave, out1, out2
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)   ' 8/8  <=> 8/8
    out1.CreateConstant 0, SampleSize / 2
    out2.CreateConstant 0, SampleSize / 2
    For Each site In TheExec.sites
        For i = 0 To (SampleSize \ 4 - 1)
            out1.ElementLite(2 * i) = inwave.ElementLite(4 * i)
            out1.ElementLite(2 * i + 1) = inwave.ElementLite(4 * i + 1)
            out2.ElementLite(2 * i) = inwave.ElementLite(4 * i + 2)
            out2.ElementLite(2 * i + 1) = inwave.ElementLite(4 * i + 3)
        Next i
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_SplitWave8")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
 
Public Sub Ze_GetSingleBitWave(ByVal inwave As DSPWave, ByRef outwave As DSPWave, ByRef outWave_fl As DSPWave)
On Error GoTo errHandler
Dim i As Long, SampleSize As Long, twos As Double, value As Double   'assume 32 bits always
Dim out1 As New DSPWave, out2 As New DSPWave
    
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.GetSingleBitWave inwave, outwave, outWave_fl
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    out1.CreateConstant 0, SampleSize, DspLong
    out2.CreateConstant 0, SampleSize, DspLong
    twos = 2# ^ 32
    For Each site In TheExec.sites
        For i = 0 To (SampleSize - 1)
        value = inwave.ElementLite(i)
        If value < 0 Then value = value + twos
            out1.ElementLite(i) = Fix(value / 2# ^ 16)
            out2.ElementLite(i) = value - out1.ElementLite(i) * (2# ^ 16)
        Next i
        outwave = out1.BitwiseOr(out2)
    Next
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_GetSingleBitWave")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_GetSingleBitWave_Jtag(ByVal inwave As DSPWave, ByRef outwave As DSPWave)
On Error GoTo errHandler
Dim i As Long, j As Long, SampleSize As Long
    
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.GetSingleBitWave_Jtag inwave, outwave
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    outwave.CreateConstant 0, SampleSize / 2
    For Each site In TheExec.sites
        For i = 0 To SampleSize \ 32 - 1
            For j = 0 To 15
                outwave.ElementLite(16 * i + j) = inwave.ElementLite(32 * i + j)
            Next
        Next
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_GetSingleBitWave_Jtag")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_GetDspWaveExMask(ByVal inwave As DSPWave, ByRef outwave As DSPWave, ByVal dicMask As DSPWave)
On Error GoTo errHandler
Dim iBit As Long, SampleSize As Long
    
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.GetDspWaveExMask inwave, outwave, dicMask
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    outwave.CreateConstant 0, SampleSize
    For Each site In TheExec.sites
        outwave = inwave.ConvertDataTypeTo(DspDouble).Multiply(dicMask)
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_GetDspWaveExMask")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_CalOneInAverage(ByVal inwave As DSPWave, ByRef result As Double, ByVal aveSamples As SiteLong)
On Error GoTo errHandler
Dim SampleSize As Long
    
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.CalOneInAverage inwave, result, aveSamples
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    For Each site In TheExec.sites
        result = inwave.ConvertDataTypeTo(DspLong).CountElements(EqualTo, 1) / aveSamples
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_CalOneInAverage")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_TwoDspWaveCompare(ByVal inWave1 As DSPWave, ByVal inWave2 As DSPWave, ByRef result As SiteLong)
On Error GoTo errHandler
Dim comp As New DSPWave
    
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.TwoDspWaveCompare inWave1, inWave2, result
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    For Each site In TheExec.sites
        comp = inWave1.ConvertDataTypeTo(DspLong).LogicalCompare(EqualTo, inWave2.ConvertDataTypeTo(DspLong))
        result = comp.CountElements(EqualTo, 0)
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_TwoDspWaveCompare")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Ze_XExpandDblBit(ByVal inwave As DSPWave, ByRef outwave As DSPWave)
On Error GoTo errHandler
Dim iBit As Long, SampleSize As Long, tmp1Dspwave As New DSPWave, tmp2Dspwave As New DSPWave
Dim appSize As Long
Dim appWave As New DSPWave
    
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    rundsp.XExpandDblBit inwave, outwave
    'Call RunDspReSet
Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    appSize = SampleSize Mod 16
    If appSize <> 0 Then
        appWave.CreateConstant 0, appSize, inwave.DataType
        For Each site In TheExec.sites
          appWave = inwave.Select(SampleSize - appSize, 1, appSize).Copy
        Next
    End If
    outwave.CreateConstant 0, SampleSize * 2
    tmp2Dspwave.CreateConstant 0, (SampleSize \ 16) * 2, DspLong
    For Each site In TheExec.sites
        tmp1Dspwave = inwave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
        For iBit = 0 To SampleSize \ 16 - 1
             tmp2Dspwave.Element(2 * iBit) = tmp1Dspwave.Element(iBit)
             tmp2Dspwave.Element(2 * iBit + 1) = tmp1Dspwave.Element(iBit)
        Next
        outwave = tmp2Dspwave.ConvertStreamTo(tldspSerial, 16, 0, Bit0IsMsb)
        If appSize <> 0 Then outwave = outwave.Concatenate(appWave).Concatenate(appWave)
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_XExpandDblBit")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Public Sub Ze_DsscValueOperation(ByVal InWave As DSPWave, ByVal fieldSt As DSPWave, ByVal fieldSp As DSPWave, ByVal msbfirst As DSPWave, ByRef outwave As DSPWave)
'If Not RunAtDsp Then GoTo s_SERIAL
'     'Call RunDspSet
'     rundsp.DsscValueOperation InWave, fieldSt, fieldSp, msbfirst, outwave
'     'Call RunDspReSet
'Exit Sub
's_SERIAL:
'    Dim iField As Long, fieldCnt As Long
'    fieldCnt = GlbUtility.GetSampleSize(fieldSt)
'    Dim stBit As Long, spBit As Long, iBit As Long, iibit As Long
'    outwave.CreateConstant 0, fieldCnt
'    outwave = outwave.ConvertDataTypeTo(DspDouble)
'    InWave = InWave.ConvertDataTypeTo(DspLong)
'            For iField = 0 To fieldCnt - 1
'                 stBit = fieldSt.ElementLite(iField)
'                 spBit = fieldSp.ElementLite(iField)
'                 iBit = spBit - stBit + 1
'                 If (iBit < 32) Then
'                         If msbfirst.ElementLite(iField) = 1 Then
'                                For iibit = iBit - 1 To 0 Step -1
'                                If InWave.ElementLite(stBit + iibit) = 1 Then outwave.ElementLite(iField) = outwave.ElementLite(iField) + 2# ^ iibit
'                                Next
'                        Else
'                                 For iibit = 0 To iBit - 1 Step 1
'                                 If InWave.ElementLite(stBit + iibit) = 1 Then outwave.ElementLite(iField) = outwave.ElementLite(iField) + 2# ^ (iBit - iibit - 1)
'                                Next
'                        End If
'                End If
'            Next
'End Sub
'Public Sub Ze_CombineShortWave(ByVal inwave As DSPWave, ByVal waveIdx As DSPWave, ByRef outwave As DSPWave, Optional ByRef dicRowNum As DSPWave, Optional DoubleBit As Boolean = True)
Public Sub Ze_CombineShortWave(ByVal inwave As DSPWave, ByVal waveIdx As DSPWave, ByVal segWave, ByRef outwave As DSPWave, Optional ByRef dicRowNum As DSPWave, Optional doubleBit As Boolean = True)
On Error GoTo errHandler
Dim iBit As Long, SampleSize As Long, iExactBit As Long, tmpwave As New DSPWave, iloc As Long, ikey As Long, tmpdicRowNum As New Dictionary
Dim bitwidth As Long: bitwidth = IIf(doubleBit, 16, 32)
Dim vKey As Variant
    
    Set dicRowNum = New DSPWave
    If Not RunAtDsp Then GoTo s_SERIAL
    'Call RunDspSet
    Dim sbitwidth As New SiteLong: sbitwidth = bitwidth
    'rundsp.CombineShortWave inWave, waveIdx, outwave, dicRowNum, sbitwidth
    rundsp.CombineShortWave inwave, waveIdx, segWave, outwave, dicRowNum, sbitwidth
    'Call RunDspReSet
    Exit Sub
s_SERIAL:
    SampleSize = GlbUtility.GetSampleSize(inwave)
    ikey = 0
    For Each site In TheExec.sites
        tmpwave = waveIdx.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, bitwidth, 0)
        outwave.CreateConstant 0, tmpwave.CountElements(GreaterThanOrEqualTo, 1) * bitwidth
        iExactBit = 0
        For iBit = 0 To SampleSize - 1
            iloc = iBit \ bitwidth
            If tmpwave.ElementLite(iloc) > 1 Then
                outwave.ElementLite(iExactBit) = inwave.ElementLite(iBit)
                iExactBit = iExactBit + 1
                If Not tmpdicRowNum.Exists(iloc) Then
                    tmpdicRowNum.Add iloc, ikey
                    ikey = ikey + 1
                End If
            End If
        Next
    Next
    dicRowNum.CreateConstant 0, tmpdicRowNum.Count
    ikey = 0
    For Each vKey In tmpdicRowNum.Keys
            dicRowNum.ElementLite(ikey) = vKey
            ikey = ikey + 1
    Next
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeDspOperation", "Ze_CombineShortWave")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Public Sub Ze_GetPartialWave(ByVal InWave As DSPWave, ByVal IdxSt As Long, ByVal size As Long, ByRef outwave As DSPWave)
'If Not RunAtDsp Then GoTo s_SERIAL
'     'Call RunDspSet
'     rundsp.GetPartialWave InWave, IdxSt, size, outwave
'     'Call RunDspReSet
'Exit Sub
's_SERIAL:
'        For Each site In TheExec.sites
'                Set outwave = InWave.Select(IdxSt, 1, size).Copy
'        Next
'End Sub
