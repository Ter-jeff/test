Attribute VB_Name = "DSP_AeFuse"
Option Explicit
Public Function SingleDoubleCal(ByVal inWave1 As DSPWave, ByVal inWave2 As DSPWave, ByRef result As Long) As Long
On Error Resume Next
Dim comp As New DSPWave
        Set comp = inWave1.LogicalCompare(EqualTo, inWave2)
        result = comp.CountElements(EqualTo, 0)
End Function

'Public Function DspwaveToSerialMsb(ByRef wave As DSPWave, ByVal parallelBits As Long) As Long
'                wave = wave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, parallelBits, 0, Bit0IsMsb)
'End Function

'Public Function DspwaveToSerialLsb(ByRef wave As DSPWave, ByVal parallelBits As Long) As Long
'                wave = wave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, parallelBits, 0, Bit0IsLsb)
'End Function

Public Function DspwaveToParallelMsb(ByVal wave As DSPWave, ByVal parallelBits As Long, ByRef outwave As DSPWave) As Long
On Error Resume Next
 Dim appSize As Long
Dim appWave As New DSPWave
     appSize = wave.SampleSize Mod parallelBits
     If appSize <> 0 Then
            appSize = parallelBits - appSize
            appWave.CreateConstant 0, appSize, wave.DataType
            wave = wave.Concatenate(appWave)
     End If
                outwave = wave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, parallelBits, 0, Bit0IsMsb)
End Function

'Public Function DspwaveToParallelLsb(ByVal wave As DSPWave, ByVal parallelBits As Long, ByRef outwave As DSPWave) As Long
' Dim appSize As Long
'     appSize = wave.SampleSize Mod parallelBits
'     If appSize <> 0 Then
'            appSize = parallelBits - appSize
'            Dim appWave As New DSPWave
'            appWave.CreateConstant 0, appSize, wave.DataType
'            wave = wave.Concatenate(appWave)
'     End If
'                outwave = wave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, parallelBits, 0, Bit0IsLsb)
'End Function

Public Function DspwaveParallel8ToParallel32(ByVal wave As DSPWave, ByRef outwave As DSPWave) As Long
On Error Resume Next
Dim i As Long, SampleSize As Long
SampleSize = wave.SampleSize
            outwave.CreateConstant 0, SampleSize / 4
                    For i = 0 To SampleSize \ 4 - 1
                        outwave.Element(i) = wave.Element(4 * i + 3) * 2 ^ 24 + wave.Element(4 * i + 2) * 2 ^ 16 + wave.Element(4 * i + 1) * 2 ^ 8 + wave.Element(4 * i)
'                        outwave.ElementLite(i) = Wave.ElementLite(4 * i + 3) * 2 ^ 24 + Wave.ElementLite(4 * i + 2) * 2 ^ 16 + Wave.ElementLite(4 * i + 1) * 2 ^ 8 + Wave.ElementLite(4 * i)
                    Next i
End Function

Public Function DspwaveParallel16ToParallel32(ByVal wave As DSPWave, ByRef outwave As DSPWave) As Long
On Error Resume Next
Dim i As Long, SampleSize As Long
    SampleSize = wave.SampleSize
    outwave.CreateConstant 0, SampleSize / 2
    For i = 0 To SampleSize \ 2 - 1
        outwave.Element(i) = wave.Element(2 * i + 1) * 2 ^ 16 + wave.Element(2 * i)
    Next i
End Function
Public Function DspwaveBlankCheck(ByRef wave As DSPWave, ByRef FBC As Long) As Long
On Error Resume Next
                FBC = wave.CountElements(GreaterThan, 0)
End Function

Public Function SplitWaveSerial32(ByVal InWave As DSPWave, ByRef out1 As DSPWave, ByRef out2 As DSPWave) As Long
On Error Resume Next
Dim i As Long, j As Long, SampleSize As Long 'assume 32 bits always
SampleSize = InWave.SampleSize
out1.CreateConstant 0, SampleSize / 2
out2.CreateConstant 0, SampleSize / 2
                    For i = 0 To (SampleSize \ 32 - 1)
                            For j = 0 To 15
                                    out1.ElementLite(16 * i + j) = InWave.ElementLite(32 * i + j)
                                    out2.ElementLite(16 * i + j) = InWave.ElementLite(32 * i + j + 16)
                            Next
                    Next i
End Function

Public Function SplitWave32(ByVal InWave As DSPWave, ByRef out1 As DSPWave, ByRef out2 As DSPWave) As Long
On Error Resume Next
Dim i As Long, SampleSize As Long, twos As Double, value As Double   'assume 32 bits always
SampleSize = InWave.SampleSize
out1.CreateConstant 0, SampleSize
out2.CreateConstant 0, SampleSize
twos = 2# ^ 32
                    For i = 0 To (SampleSize - 1)
                    value = InWave.ElementLite(i)
                    If value < 0 Then value = value + twos
                            out1.ElementLite(i) = Fix(value / 2# ^ 16)
                            out2.ElementLite(i) = value - out1.ElementLite(i) * (2# ^ 16)
                    Next i
End Function
 
Public Function SplitWave8(ByVal InWave As DSPWave, ByRef out1 As DSPWave, ByRef out2 As DSPWave) As Long
On Error Resume Next
Dim i As Long, SampleSize As Long  'assume 32 bits always
SampleSize = InWave.SampleSize    ' 8/8  <=> 8/8
out1.CreateConstant 0, SampleSize / 2
out2.CreateConstant 0, SampleSize / 2
                    For i = 0 To (SampleSize \ 4 - 1)
                            out1.ElementLite(2 * i) = InWave.ElementLite(4 * i)
                            out1.ElementLite(2 * i + 1) = InWave.ElementLite(4 * i + 1)
                            out2.ElementLite(2 * i) = InWave.ElementLite(4 * i + 2)
                            out2.ElementLite(2 * i + 1) = InWave.ElementLite(4 * i + 3)
                   Next i
End Function
 
Public Function GetSingleBitWave(ByVal InWave As DSPWave, ByRef outwave As DSPWave, ByRef outWave_fl As DSPWave) As Long
On Error Resume Next
Dim i As Long, SampleSize As Long, twos As Double, value As Double   'assume 32 bits always
Dim out1 As New DSPWave, out2 As New DSPWave
Dim convertData As New DSPWave
Dim tmp As New DSPWave

convertData = InWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
SampleSize = InWave.SampleSize

        For i = 0 To (SampleSize - 1)
            tmp = convertData.Select(32 * i, 1, 16).ConvertDataTypeTo(DspLong).COPY.BitwiseOr(convertData.Select(32 * i + 16, 1, 16).ConvertDataTypeTo(DspLong).COPY)
            out1 = out1.ConvertDataTypeTo(DspLong).Concatenate(tmp)
        Next i
   Set outWave_fl = InWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).COPY
    Set outwave = out1

'Dim i As Long, SampleSize As Long, twos As Double, value As Double   'assume 32 bits always
'Dim out1 As New DSPWave, out2 As New DSPWave
'Dim convertData As New DSPWave
'SampleSize = inwave.SampleSize
'out1.CreateConstant 0, SampleSize, DspLong
'out2.CreateConstant 0, SampleSize, DspLong

'twos = 2# ^ 32
'                    For i = 0 To (SampleSize - 1)
'                    value = inwave.ElementLite(i)
'                    If value < 0 Then value = value + twos
'                            out1.ElementLite(i) = Fix(value / 2# ^ 16)
'                            out2.ElementLite(i) = value - out1.ElementLite(i) * (2# ^ 16)
'                    Next i
'    Set outWave_fl = inwave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).Copy
'    Set outwave = out1.BitwiseOr(out2)
End Function

Public Function GetSingleBitWave_Jtag(ByVal InWave As DSPWave, ByRef outwave As DSPWave) As Long
On Error Resume Next
Dim i As Long, j As Long, SampleSize As Long
Dim tmp As New DSPWave
Dim LoopSize As Long
SampleSize = InWave.SampleSize
LoopSize = SampleSize / 32
'outwave.CreateConstant 0, SampleSize / 2

        For i = 0 To LoopSize - 1
            tmp = InWave.Select(32 * i, 1, 16).ConvertDataTypeTo(DspLong).COPY.BitwiseOr(InWave.Select(32 * i + 16, 1, 16).ConvertDataTypeTo(DspLong).COPY)
            If i = 0 Then
                outwave = tmp.COPY
            Else
                outwave = outwave.Concatenate(tmp).COPY
            End If
        Next i




'Dim i As Long, j As Long, SampleSize As Long
'SampleSize = inwave.SampleSize
'outwave.CreateConstant 0, SampleSize / 2
'                    For i = 0 To SampleSize \ 32 - 1
'                            For j = 0 To 15
'                                    outwave.ElementLite(16 * i + j) = inwave.ElementLite(32 * i + j)
'                            Next
'                    Next
End Function

Public Function GetDspWaveExMask(ByVal InWave As DSPWave, ByRef outwave As DSPWave, ByVal dicMask As DSPWave) As Long
On Error Resume Next
    Dim iBit As Long, SampleSize As Long
    SampleSize = InWave.SampleSize
    outwave.CreateConstant 0, SampleSize
            outwave = InWave.ConvertDataTypeTo(DspDouble).Multiply(dicMask)
End Function

Public Function CalOneInAverage(ByVal InWave As DSPWave, ByRef result As Double, ByVal aveSamples As Long) As Long
On Error Resume Next
  Dim SampleSize As Long
    SampleSize = InWave.SampleSize
    result = InWave.ConvertDataTypeTo(DspLong).CountElements(EqualTo, 1) / aveSamples
End Function

Public Function TwoDspWaveCompare(ByVal inWave1 As DSPWave, ByVal inWave2 As DSPWave, ByRef result As Double) As Long
On Error Resume Next
Dim comp As New DSPWave
        Set comp = inWave1.ConvertDataTypeTo(DspLong).LogicalCompare(EqualTo, inWave2.ConvertDataTypeTo(DspLong))
        result = comp.CountElements(EqualTo, 0)
End Function

Public Function XExpandDblBit(ByVal InWave As DSPWave, ByRef outwave As DSPWave) As Long
On Error Resume Next
    Dim iBit As Long, SampleSize As Long, tmp1Dspwave As New DSPWave, tmp2Dspwave As New DSPWave
Dim appSize As Long
Dim appWave As New DSPWave
    SampleSize = InWave.SampleSize
    appSize = SampleSize Mod 16
    If appSize <> 0 Then
           appWave = InWave.Select(SampleSize - appSize, 1, appSize).COPY
    End If
    outwave.CreateConstant 0, SampleSize * 2
    tmp2Dspwave.CreateConstant 0, (SampleSize \ 16) * 2, DspLong
       tmp1Dspwave = InWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
       For iBit = 0 To SampleSize \ 16 - 1
            tmp2Dspwave.Element(2 * iBit) = tmp1Dspwave.Element(iBit)
            tmp2Dspwave.Element(2 * iBit + 1) = tmp1Dspwave.Element(iBit)
       Next
       Set outwave = tmp2Dspwave.ConvertStreamTo(tldspSerial, 16, 0, Bit0IsMsb)
       If appSize <> 0 Then outwave = outwave.Concatenate(appWave).Concatenate(appWave)
End Function

Public Function XExpandBitVector(ByVal InWave As DSPWave, ByVal dealVecRpt As Long, ByRef outwave As DSPWave) As Long
On Error Resume Next
    outwave = InWave.Repeat(dealVecRpt).MatrixTranspose(dealVecRpt)
End Function

'Public Function DsscValueOperation(ByVal InWave As DSPWave, ByVal fieldSt As DSPWave, ByVal fieldSp As DSPWave, ByVal msbfirst As DSPWave, ByRef outwave As DSPWave) As Long
'    Dim iField As Long, fieldCnt As Long
'    fieldCnt = fieldSt.SampleSize
'    Dim stBit As Long, spBit As Long, iBit As Long, iibit As Long
'    InWave = InWave.ConvertDataTypeTo(DspLong)
'    outwave.CreateConstant 0, fieldCnt
'    outwave = outwave.ConvertDataTypeTo(DspDouble)
'    For iField = 0 To fieldCnt - 1
'         stBit = fieldSt.ElementLite(iField)
'         spBit = fieldSp.ElementLite(iField)
'         iBit = spBit - stBit + 1
'         If (iBit < 32) Then
'                 If msbfirst.ElementLite(iField) = 1 Then
'                        For iibit = iBit - 1 To 0 Step -1
'                        If InWave.ElementLite(stBit + iibit) = 1 Then outwave.ElementLite(iField) = outwave.ElementLite(iField) + 2# ^ iibit
'                        Next
'                Else
'                         For iibit = 0 To iBit - 1 Step 1
'                         If InWave.ElementLite(stBit + iibit) = 1 Then outwave.ElementLite(iField) = outwave.ElementLite(iField) + 2# ^ (iBit - iibit - 1)
'                        Next
'                End If
'        End If
'    Next
'End Function

Public Function CombineShortWave(ByVal InWave As DSPWave, ByVal waveIdx As DSPWave, ByVal segWave As DSPWave, ByRef outwave As DSPWave, ByRef dicRowNum As DSPWave, ByVal bitwidth As Long) As Long
On Error Resume Next
    Dim iBit As Long, SampleSize As Long, iExactBit As Long, tmpwave32 As New DSPWave, iloc As Long, ikey As Long, rowSize As Long, preLoc As Long

    ''-------------------------------------
    Dim m_SegFlag() As Long
    Dim offset As Long
    Dim cnt As Long
    SampleSize = InWave.SampleSize
    tmpwave32 = waveIdx.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, bitwidth, 0)
    rowSize = tmpwave32.CountElements(GreaterThanOrEqualTo, 1)
    If rowSize = 0 Then Exit Function
    outwave.CreateConstant 0, rowSize * bitwidth
    dicRowNum.CreateConstant 0, rowSize
    iExactBit = 0
    ikey = 0: preLoc = 0
    cnt = 0
    m_SegFlag = segWave.data
    For iBit = 0 To segWave.SampleSize - 1
        offset = cnt * bitwidth
        If m_SegFlag(iBit) = 1 Then
            outwave.Select(offset, 1, bitwidth).Replace InWave.Select(bitwidth * iBit, 1, bitwidth)
            dicRowNum.ElementLite(cnt) = iBit
            cnt = cnt + 1
        End If
    Next iBit
    ''-----------------------


''      For iBit = 0 To SampleSize - 1
''            iloc = iBit \ bitwidth
''            If tmpwave32.ElementLite(iloc) > 0 Then
''                    'outWave.ElementLite(iExactBit) = inWave.ElementLite(iBit)
''                    iExactBit = iExactBit + 1
''                    If ikey < rowSize Then
''                                    If preLoc <> iloc Then
''                                            dicRowNum.ElementLite(ikey) = iloc
''                                            ikey = ikey + 1
''                                            preLoc = iloc
''                                    End If
''                   End If
''            End If
''      Next
End Function
''Public Function CombineShortWave(ByVal inwave As DSPWave, ByVal waveIdx As DSPWave, ByRef outwave As DSPWave, ByRef dicRowNum As DSPWave, ByVal bitwidth As Long) As Long
''    Dim iBit As Long, SampleSize As Long, iExactBit As Long, tmpwave32 As New DSPWave, iloc As Long, ikey As Long, rowSize As Long, preLoc As Long
''    SampleSize = inwave.SampleSize
''    tmpwave32 = waveIdx.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, bitwidth, 0)
''    rowSize = tmpwave32.CountElements(GreaterThanOrEqualTo, 1)
''    outwave.CreateConstant 0, rowSize * bitwidth
''    dicRowNum.CreateConstant 0, rowSize
''    iExactBit = 0
''    ikey = 0: preLoc = 0
''      For iBit = 0 To SampleSize - 1
''            iloc = iBit \ bitwidth
''            If tmpwave32.ElementLite(iloc) > 1 Then
''                    outwave.ElementLite(iExactBit) = inwave.ElementLite(iBit)
''                    iExactBit = iExactBit + 1
''                    If ikey < rowSize Then
''                                    If preLoc <> iloc Then
''                                            dicRowNum.ElementLite(ikey) = iloc
''                                            ikey = ikey + 1
''                                            preLoc = iloc
''                                    End If
''                   End If
''            End If
''      Next
''End Function
'Public Sub GetPartialWave(ByVal InWave As DSPWave, ByVal IdxSt As Long, ByVal size As Long, ByRef outwave As DSPWave)
'                Set outwave = InWave.Select(IdxSt, 1, size).Copy
'End Sub

Public Function CompareBlkChkData(ByVal inWave1 As DSPWave, ByVal inWave2 As DSPWave, ByRef m_Result As Long) As Long
On Error Resume Next
    Dim outwave As New DSPWave

    outwave.CreateConstant 1, inWave1.SampleSize, DspLong

    outwave = inWave1.ConvertDataTypeTo(DspLong).LogicalCompare(NotEqualTo, inWave2.ConvertDataTypeTo(DspLong))
    m_Result = outwave.CalcSum
End Function

Public Function GetDspWaveFailbits(ByVal InWave As DSPWave, ByVal MaskWave As DSPWave, ByRef m_Result As Long) As Long
On Error Resume Next
    Dim tmpWave As New DSPWave

    tmpWave.CreateConstant 1, InWave.SampleSize, DspLong

    tmpWave = InWave.ConvertDataTypeTo(DspLong).bitwiseand(MaskWave)
    m_Result = tmpWave.CountElements(EqualTo, 1)
End Function


Public Function CompareWRData(ByVal inWave1 As DSPWave, ByVal inWave2 As DSPWave, ByVal MaskWave As DSPWave, ByRef m_Result As Long, ByVal Rvenable As Boolean) As Long
On Error Resume Next
    Dim outwave As New DSPWave
    Dim tmpWave As New DSPWave
    
    outwave.CreateConstant 1, inWave1.SampleSize, DspLong
    
    If Rvenable Then
        outwave = inWave1.ConvertDataTypeTo(DspLong).LogicalCompare(NotEqualTo, inWave2.ConvertDataTypeTo(DspLong))
    Else
        tmpWave = inWave2.ConvertDataTypeTo(DspLong).bitwiseand(MaskWave)
        outwave = inWave1.ConvertDataTypeTo(DspLong).LogicalCompare(NotEqualTo, tmpWave.ConvertDataTypeTo(DspLong))
    End If
    m_Result = outwave.CalcSum

    
End Function
Public Function ConvertToDecimal(ByVal InWave As DSPWave, ByVal fieldSt As DSPWave, ByVal fieldSp As DSPWave, ByVal msbfirst As DSPWave, ByRef outwave As DSPWave) As Long
On Error Resume Next
    Dim iField As Long, fieldCnt As Long
    Dim stBit As Long, spBit As Long, iBit As Long, iibit As Long
    Dim Wave2 As New DSPWave

    fieldCnt = fieldSt.SampleSize
    InWave = InWave.ConvertDataTypeTo(DspLong)
    outwave.CreateConstant 0, fieldCnt, DspLong
    For iField = 0 To fieldCnt - 1
         stBit = fieldSt.ElementLite(iField)
         spBit = fieldSp.ElementLite(iField)
         iBit = spBit - stBit + 1
         If (iBit < 32) Then
                 If msbfirst.ElementLite(iField) = 1 Then
                        Wave2 = InWave.Select(stBit, 1, iBit).COPY
                        outwave.Element(iField) = Wave2.ConvertStreamTo(tldspParallel, iBit, 0, Bit0IsMsb).Element(0)
                Else
                        Wave2 = InWave.Select(stBit, 1, iBit).COPY
                        outwave.Element(iField) = Wave2.ConvertStreamTo(tldspParallel, iBit, 0, Bit0IsLsb).Element(0)
                End If
        End If
    Next
End Function

Public Function ProcessCapWaveAndCheck(ByRef InWave As DSPWave, _
                                       ByVal MaskWave As DSPWave, _
                                       ByVal parallelSize As Long, _
                                       ByVal accessMode As Long, _
                                       ByVal DoubleBits As Long, _
                                       ByVal JtagMsbFirst As Boolean, _
                                       ByRef ParallelWave As DSPWave, _
                                       ByRef SerialWave As DSPWave, _
                                       ByRef SerialWave_FL As DSPWave, _
                                       ByRef blankresult As Long, _
                                       ByVal fieldStart As DSPWave, _
                                       ByVal fieldStop As DSPWave, _
                                       ByVal msbF As DSPWave, _
                                       ByRef decimalOut As DSPWave, _
                                       ByVal Rvenable As Boolean, _
                                       ByVal rvWaveIdx As DSPWave, _
                                       ByVal segWave As DSPWave, _
                                       ByRef rvOutWave As DSPWave, _
                                       ByVal TestedWaveCheck As Boolean, _
                                       ByVal Testedstagemask As DSPWave, _
                                       ByRef capwave_count As Long, _
                                       ByVal CapPinCnt As Long) As Long

On Error Resume Next
Dim idxWave As New DSPWave
Dim waveSize As Long
Dim outwave As New DSPWave
Dim tmpWave As New DSPWave

blankresult = False

''accessMode  = 0 -> Direct Access Mode(capwave : parallel bit on every cycle)
''accessMode <> 0 -> Jtag Mode         (capwave : serial bit on every cycle)
    If accessMode = 0 Then
        If CapPinCnt < parallelSize Then InWave = InWave.ConvertDataTypeTo(DspLong).ConvertNumFormatTo(0, CapPinCnt)

        ''arrange parallel bit size to 32bit
        ''RF device is small, so the capture parallel size would be 8/16 possible with every cycle.
        If parallelSize = 32 Then
            If CapPinCnt = 16 Then
                Call DspwaveParallel16ToParallel32(InWave, ParallelWave)
            Else
            ParallelWave = InWave.COPY
            End If
        ElseIf parallelSize = 8 Then
            Call DspwaveParallel8ToParallel32(InWave, ParallelWave)
        End If

        ''Parallel wave converts to serial wave with 2-bit/1-bit case.
        If DoubleBits = 2 Then
            GetSingleBitWave ParallelWave, SerialWave, SerialWave_FL
            'DspwaveToSerialMsb SerialWave, 16
        Else
            Set SerialWave = ParallelWave
            SerialWave = SerialWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
            'DspwaveToSerialMsb SerialWave, 32
        End If
    Else
        If DoubleBits = 2 Then
            GetSingleBitWave_Jtag InWave, SerialWave
        Else
            Set SerialWave = InWave
        End If
        waveSize = SerialWave.SampleSize
        If JtagMsbFirst Then idxWave.CreateRamp waveSize - 1, -1, waveSize, DspLong
        If JtagMsbFirst Then SerialWave = SerialWave.Lookup(idxWave)

    End If


    If TestedWaveCheck Then
        Call GetDspWaveExMask(SerialWave, outwave, Testedstagemask)
        Call DspwaveBlankCheck(outwave, capwave_count)
        Set outwave = New DSPWave
    End If

    If MaskWave.CalcSum = 0 Then GoTo skip

    ''MaskWave(current stage wave) masks with serial wave to filter the read wave for current job.
    Call GetDspWaveExMask(SerialWave, outwave, MaskWave)
    ''Check the wave blank for current job after the filter process.
    Call DspwaveBlankCheck(outwave, blankresult)

skip:
   ''converts the serail wave to decimal data for every field.
    ConvertToDecimal SerialWave, fieldStart, fieldStop, msbF, decimalOut
    'if Rvenable Then CombineShortWave outwave, rvWaveIdx, rvOutWave, tmpWave, 16
    If Rvenable Then CombineShortWave outwave, rvWaveIdx, segWave, rvOutWave, tmpWave, 32
End Function

