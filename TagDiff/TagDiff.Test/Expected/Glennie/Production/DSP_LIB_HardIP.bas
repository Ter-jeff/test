Attribute VB_Name = "DSP_LIB_HardIP"

Option Explicit

' This module should be used only for DSP Procedure code.  Functions in this
' module will be available to be called to perform DSP in all DSP modes.
' Additional modules may be added as needed (all starting with "DSP_").
'
' The required signature for a DSP Procedure is:
'
' Public Function FuncName(<arglist>) as Long
'   where <arglist> is any list of arguments supported by DSP code.
'
' See online help for supported types and other restrictions.

Public Function Measure_DutyCycle(ByVal Wave1 As DSPWave, ByRef dutycycle As Double, ByRef dspStatus As Long) As Long

''Dim Rj As Double
''Dim DDj As Double
''Dim MeasuredUI As Double

Dim RJ As Double
Dim DDJ As Double
Dim MeasuredUI As Double



Call Wave1.MeasureJitter(RJ, DDJ, MeasuredUI, dspStatus)

dutycycle = 100# * ((MeasuredUI - DDJ) / (2 * MeasuredUI))


TheHdw.Digital.Jitter.FileExport Wave1, ".\LPDPADCJitter1"




End Function

Public Function trim_calc(ByVal capWave As DSPWave, ByRef outwave As DSPWave) As Long
Dim i As Integer

    'calculate trim code
'        OutWave.Element(0) = 12
    outwave.CreateConstant 0, 5
    
        For i = 0 To 4 'CapWave.CountElements
       
        If (capWave.Element(i) < 274.42) Then
         outwave.Element(i) = 0
        Else
         outwave.Element(i) = (73.7579 - 10.2598 * Sqr(capWave.Element(i) - 274.42)) ' update 20130916 ccwu
        End If
        
        'OutWave.Element(i) = (73.134 - 10.153 * Sqr(CapWave.Element(i) - 272.32))  ' 25C
        
        'OutWave.Element(i) = (73.7579 - 10.2598 * Sqr(CapWave.Element(i) - 274.42)) ' update 20130916 ccwu
        'OutWave.Element(i) = 15
        'OutWave.Element(i) = (92.8819 - 6.4957 * Sqr(CapWave.Element(i) - 786.95)) ' 125C
        
            If outwave.Element(i) > 31 Then
                outwave.Element(i) = 31
            ElseIf outwave.Element(i) < 0 Then
                outwave.Element(i) = 0
            End If
        
        Next i
    
End Function


Public Function trim_conveter(ByVal InWave As DSPWave, ByRef outwave As DSPWave) As Long
Dim i As Long, j As Long
Dim in_temp As Long

    outwave.CreateConstant 0, 3
    
     outwave.Element(0) = InWave.Element(2)
     outwave.Element(1) = InWave.Element(3)
     outwave.Element(2) = InWave.Element(4)
    
''    For j = 0 To 1
''      in_temp = inWave.Element(j)
''        For i = 0 To 4
''          OutWave.Element(j) = in_temp Mod 2
''          in_temp = Int(in_temp / 2)
''        Next i
''    Next j
End Function


Public Function MeasureJitter_US10G(ByVal wave As DSPWave, DDJ As Double, RJ As Double, Meas_UI As Double, dspStatus As Long, ByRef dutycycle As Double, ByRef freq As Double) As Long
'used to be MeasureJitter_vb
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim MeasuredPeriod As Double
   ' Dim dspStatus As Long
   ' Dim dspStatus1 As Long
    Dim UI As Double

    Call wave.SerialMeasureJitter(RJ, DDJ, Meas_UI, dspStatus)
    Call wave.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
    End Function

Public Function Measure_Eye_US10G(ByVal wave As DSPWave, UI As Double, RJ As Double, DDJ As Double, RiseTime As Double, _
    FallTime As Double, EarlyLow As Double, LateLow As Double, EarlyMid As Double, LateMid As Double, EarlyHigh As Double, _
    LateHigh As Double, dspStatus As Long) As Long
'used to be analyzeEye_vb
    Call wave.SerialMeasureEye(RJ, DDJ, UI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LateMid, EarlyHigh, LateHigh, dspStatus)
End Function

Public Function Measure_DutyCycle_Freq_US10G(ByVal wave As DSPWave, dutycycle As Double, freq As Double) As Long
    Dim RJ        As Double
    Dim DDJ       As Double
    Dim UI        As Double
    Dim dspStatus As Long

    Call wave.SerialMeasureJitter(RJ, DDJ, UI, dspStatus)

    dutycycle = 100# * (UI - DDJ) / (2 * UI)
    freq = 1 / (2 * UI)
    
End Function

Public Function LoopEyeMeas(ByVal wave As DSPWave, RJ As Double, DDJ As Double, Tj As Double, measUI As Double, Tr As Double, _
            Tf As Double, Eye50 As Double, Eye20 As Double, Eye80 As Double, retStatus As Long) As Long
        
    Dim EarlyLow As Double
    Dim LateLow As Double
    Dim EarlyMid As Double
    Dim LateMid As Double
    Dim EarlyHi As Double
    Dim LateHi As Double
    Dim RiseTime As Double
    Dim FallTime As Double
        
    Call wave.MeasureEye(RJ, DDJ, measUI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LateMid, EarlyHi, LateHi, retStatus)

    Call TheHdw.Digital.Jitter.FileExport(wave, "Raw Eye Data")
    
    Eye20 = EarlyLow + measUI - LateLow
    Eye50 = EarlyMid + measUI - LateMid
    Eye80 = EarlyHi + measUI - LateHi
    
End Function

Public Function duty_freq_meas(ByVal Captured As DSPWave, ByRef dutycycle As Double, ByRef freq As Double) As Long
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim MeasuredPeriod As Double
    Dim dspStatus As Long
    
    
    Call Captured.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
End Function


Public Function pulseMeas(ByVal wave As DSPWave, PWHigh As Double, PWLow As Double, period As Double, Status As Long) As Long
        

        
    Call wave.measuretime(PWHigh, PWLow, period, Status)
    
End Function
    
Public Function LoopJitterMeas(ByVal wave As DSPWave, RJ As Double, DDJ As Double, measUI As Double, retStatus As Long, ByRef dutycycle As Double, ByRef freq As Double) As Long
        
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim MeasuredPeriod As Double
    Dim dspStatus As Long
        
    Call wave.MeasureJitter(RJ, DDJ, measUI, retStatus)

    Call TheHdw.Digital.Jitter.FileExport(wave, "Raw Jitter Data")
    
    Call wave.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
End Function

Public Function DDR_LoopJitterMeas(ByVal wave As DSPWave, RJ As Double, DDJ As Double, measUI As Double, retStatus As Long, ByRef dutycycle As Double, ByRef freq As Double, _
                                                     ByRef PWHigh As Double, ByRef PWLow As Double, ByRef MeasuredPeriod As Double) As Long
        
''    Dim PWHigh As Double
''    Dim PWLow As Double
''    Dim MeasuredPeriod As Double
    Dim dspStatus As Long
        
''    Call wave.MeasureJitter(RJ, DDJ, measUI, retStatus)
''
''    Call TheHdw.Digital.Jitter.FileExport(wave, "Raw Jitter Data")
    
    Call wave.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
End Function

Public Function Measure_Eye(ByVal Wave1 As DSPWave, ByRef RJ As Double, ByRef DDJ As Double, _
        ByRef MeasuredUI As Double, ByRef MidWidth As Double, ByRef LowWidth As Double, ByRef HighWidth As Double, _
        ByRef RiseTime As Double, ByRef FallTime As Double, ByRef dspStatus As Long, ByVal v As Double, ByRef EarlyLow As Double, ByRef LateLow As Double, ByRef EarlyHigh As Double, ByRef LateHigh As Double) As Long

    Dim EarlyMid As Double
    Dim LateMid As Double

    
    Call Wave1.MeasureEye(RJ, DDJ, MeasuredUI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LateMid, EarlyHigh, LateHigh, dspStatus)
    ''wave1.MeasureJitter Rj, DDj, MeasuredUI, dspStatus
    MidWidth = EarlyMid + MeasuredUI - LateMid
    LowWidth = EarlyLow + MeasuredUI - LateLow
    HighWidth = EarlyHigh + MeasuredUI - LateHigh

'    thehdw.Digital.Jitter.FileExport wave1, ".\AMPAPLLJitter_" & v & "V.txt"

    Debug.Print
End Function

Public Function Measure_Jitter( _
        ByVal CapDSPWave As DSPWave, _
        ByRef resultDspWave As DSPWave, _
        ByRef dspStatus As Long) As Long
        '
        Dim calcDspWave As New DSPWave
        Dim RJ As Double
        Dim DDJ As Double
        Dim MeasuredUI As Double
        
        'Create results DSPwave for more compact storage of test results
        Set resultDspWave = New DSPWave
        resultDspWave.CreateConstant 0, 3
        '
        calcDspWave = CapDSPWave.COPY
        'calcDspWave.Plot "Jitter Test:"
        Call calcDspWave.MeasureJitter(RJ, DDJ, MeasuredUI, dspStatus)
        '
        resultDspWave.Element(0) = RJ
        resultDspWave.Element(1) = DDJ
        resultDspWave.Element(2) = MeasuredUI

        
End Function

Public Function MeasureJitter_vb(ByVal wave As DSPWave, DDJ As Double, RJ As Double, Meas_UI As Double, dspStatus As Long, ByRef dutycycle As Double, ByRef freq As Double) As Long
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim MeasuredPeriod As Double
   ' Dim dspStatus As Long
   ' Dim dspStatus1 As Long
    Dim UI As Double

    Call wave.SerialMeasureJitter(RJ, DDJ, Meas_UI, dspStatus)
    Call wave.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
    End Function

Public Function analyzeEye_vb(ByVal wave As DSPWave, UI As Double, RJ As Double, DDJ As Double, RiseTime As Double, _
    FallTime As Double, EarlyLow As Double, LateLow As Double, EarlyMid As Double, LateMid As Double, EarlyHigh As Double, _
    LateHigh As Double, dspStatus As Long) As Long
    
    Call wave.SerialMeasureEye(RJ, DDJ, UI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LateMid, EarlyHigh, LateHigh, dspStatus)
End Function

Public Function Measuredutycycle(ByVal wave As DSPWave, dutycycle As Double, freq As Double) As Long
 
    Dim RJ        As Double
    Dim DDJ       As Double
    Dim UI        As Double
    Dim dspStatus As Long

    Call wave.SerialMeasureJitter(RJ, DDJ, UI, dspStatus)

    dutycycle = 100# * (UI - DDJ) / (2 * UI)
    freq = 1 / (2 * UI)
    
End Function

Public Function BitWf2Arry(ByVal InWf As DSPWave, ByVal WrdWdth As Integer, ByRef NoOfSamples As Long, ByRef DataWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Convert captured (serial) bit stream to data waveform, Assume LSB->MSB in the bit stream (reversed
    ''''        order may be easily accommodated by adding a switch in the argument list)
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        BitWf2Arry is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        InWf          : DSP Wave (serial) to be converted
    ''''        WrdWdth  : number of bits per word
    ''''        NoOfSamples    : number of samples found in the bit stream
    ''''        DataWf         : converted (parallel) DSP Wave
    ''''
          
    NoOfSamples = InWf.SampleSize
    
    If NoOfSamples Mod WrdWdth <> 0 Then
         Debug.Print vbNewLine & "Bit stream wave size not integer times of the word width." _
            & " Waveform will Be truncated" & vbNewLine
    End If
    
    DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
    NoOfSamples = DataWf.SampleSize

End Function



Public Function DdrSwpWf2Arry(ByVal InWf As DSPWave, ByVal NoOfBists As Long, ByVal EyeStrobes As Long, _
    ByVal NoOfMdlls As Long, ByVal MdllWrdWidth As Long, ByRef EyeWf As DSPWave, ByRef MdllWf As DSPWave) As Long
    
    ''''--------------------------------------------------------------------------------------------------
    ''''    Convert captured DDR sweep bit stream waveform to data array, ssume LSB->MSB in the bit stream.
    ''''        order may be easily accommodated by adding a switch in the argument list)
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        DdrSwpWf2Arry is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        InWf            : DSP Wave (serial) to be converted
    ''''        NoOfBists       : DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        EyeStrobes      : Number of strobes for each eye in this sweep
    ''''        NoOfMdlls       : Number of MDLL cal codes to be captured.
    ''''        MdllWrdWidth    : Number of bits in each MDLL cal code
    ''''        EyeWf           : Converted eye diagram waveform to be passed back
    ''''        MdllWf          : Converted MDLL code array in waveform to be passed back
        
    Dim NoOfSamples As Long
    Dim SegWf As New DSPWave, tmpwf As New DSPWave
    Dim BistIdx As Long, SegIdx   As Long
    Dim NoOfSegs As Long, SegWidth As Long
    Dim BistsPerMdll As Long    '''' MDLL is per lane, while BIST may be per byte.
                                '''' MDLL code will be captured per lane, followed by all the eyes in that lane
    
    '''' Not all loopback involve MDLL code
    If NoOfMdlls <> 0 Then
        If NoOfBists Mod NoOfMdlls <> 0 Then
            Debug.Print vbNewLine & "No of BISTs must be multiples of Number of MDLLs!" & vbNewLine
        End If
        NoOfSegs = NoOfMdlls
        BistsPerMdll = NoOfBists / NoOfMdlls
        SegWidth = BistsPerMdll * EyeStrobes + MdllWrdWidth
    Else
        BistsPerMdll = NoOfBists
        SegWidth = EyeStrobes * NoOfBists
        NoOfSegs = 1
    End If
        
    EyeWf.CreateConstant 0, EyeStrobes * NoOfBists, DspLong
    MdllWf.CreateConstant 0, NoOfMdlls, DspLong
      
    '''' construct the eye diagram and MDLL waveforms from the captured bit stream
    For SegIdx = 0 To NoOfSegs - 1
        SegWf = InWf.Select(SegIdx * SegWidth, 1, SegWidth).COPY
        If MdllWrdWidth <> 0 Then
            tmpwf = SegWf.Select(0, 1, MdllWrdWidth).COPY
            MdllWf.Element(SegIdx) = tmpwf.ConvertStreamTo(tldspParallel, MdllWrdWidth, 0, Bit0IsMsb).Element(0)
            tmpwf.Clear
        End If
        
        For BistIdx = 0 To BistsPerMdll - 1
            Dim st0 As Long, st1 As Long
            st0 = EyeStrobes * BistIdx
            st1 = SegIdx * (SegWidth - MdllWrdWidth) + st0
            tmpwf = SegWf.Select(MdllWrdWidth + st0, 1, EyeStrobes).COPY
            EyeWf.Select(st1, 1, EyeStrobes).Replace tmpwf
            tmpwf.Clear
        Next BistIdx
    Next SegIdx
    
End Function
Public Function FindMaxEyeWidth(ByVal SwqEyeWf As DSPWave, ByVal SwsEyeWf As DSPWave, _
                        ByVal NoOfBists As Integer, ByRef EyeWidthWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Search for the eyewidth corresponding to the largest eye in the presense of one more eyes.
    ''''    An eye is defined as a continuous "1's"
    ''''    Based on the test methodologies used for TMA, there are 2 sweeps from the center of the UI, the first
    ''''        towards left, and second right, Swq and Sws respectively. The function would stitch them together by
    ''''        reversing the Swq waveform, and concatenating with the Sws wave.
    ''''
    ''''    The resulting eye diagram is from the left to right covering the entire UI.

    
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016

    ''''    Usage
    ''''        FindMaxEyeWidth is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        SwqEyeWf,SwsEyeWf   : Eye diagram waveforms, all BISTs concatenated.
    ''''        NoOfBists           : DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        EyeWidthWf          : Array of the eyewidths found, in DSP wave format, to be passed back
    
    ''''    rev 1, by Tim, 2/26/2016
    ''''
    ''''--------------------------------------------------------------------------------------------------
    ''''           New add to fix out of index error, exit for not an good method. Need to change it.
    ''''--------------------------------------------------------------------------------------------------

    Dim WholeEyeWf As New DSPWave
    Dim NoOfSamples As Long, EyeStrobes As Long
    Dim EyeWidth As Long
    Dim BistIdx As Long, idx As Long, idx2 As Long
    Dim MaxEye As Long
    EyeWidthWf.CreateConstant 0, NoOfBists, DspLong
    
    NoOfSamples = SwqEyeWf.SampleSize
    If NoOfSamples <> SwsEyeWf.SampleSize Then
             Debug.Print vbNewLine & "The lengths of Swq and Sws eye sweep waveform not consistent!" & vbNewLine
    End If
    
    EyeStrobes = NoOfSamples / NoOfBists
    
    For BistIdx = 0 To NoOfBists - 1
        WholeEyeWf = SwqEyeWf.Select(BistIdx * EyeStrobes, 1, EyeStrobes).COPY   '''' dummy operation, to allocate element
        
        '''' SWQ is from right to left, starting the middle. Reversing the eye diagram
        For idx = 0 To EyeStrobes - 1
            'idx2 = (BistIdx + 1) * EyeStrobes - 1 - idx
            WholeEyeWf.Element(idx) = SwqEyeWf.Element(idx)
        Next idx
        
        '''' stitch the reversed SWQ eye diagram to the SWK one
        WholeEyeWf = WholeEyeWf.Concatenate(SwsEyeWf.Select(BistIdx * EyeStrobes, 1, EyeStrobes).COPY)
        
        MaxEye = 0
        Dim EyeResultFlag As Boolean
        
        
        
        '''' finding the maximum eyewidth
        For idx = 0 To 2 * EyeStrobes - 1
            EyeResultFlag = False
            If WholeEyeWf.Element(idx) = 1 Then '''' starting a new eye
                EyeWidth = 0
                Do
                    EyeWidth = EyeWidth + 1
                    idx = idx + 1
               
                If idx >= 2 * EyeStrobes Then
                    EyeResultFlag = True
                ElseIf WholeEyeWf.Element(idx) = 0 Then
                    EyeResultFlag = True
                Else 'Do nothing '20230601
                End If
               
               
               'Loop 'Until WholeEyeWf.Element(idx) = 0
                        
              '  If WholeEyeWf.Element(idx) = 0 Then EyeResultFlag = True
                
                
                'Loop Until idx >= 2 * EyeStrobes - 1 Or WholeEyeWf.Element(idx) = 0
                Loop Until EyeResultFlag = True
                
                    '''' >= workaround for not being able to escape the loop while maintain the wave index within the range
                    ''''    in the condition
                    ''' The next statement separately including the last strobe
                If idx = 2 * EyeStrobes - 1 Then EyeWidth = EyeWidth + WholeEyeWf.Element(idx)
                If MaxEye < EyeWidth Then MaxEye = EyeWidth
            End If
        Next idx
        
        EyeWidthWf.Element(BistIdx) = MaxEye
        
    Next BistIdx
    

End Function

Public Function FlexibleBitWf2Arry(ByVal InWf As DSPWave, ByVal StartIndex As Long, ByVal WrdWdth As Long, ByRef DataWf As DSPWave, ByRef DataWf_Binary As DSPWave) As Long
    
    'FlexibleBitWf2Arry
    ''''--------------------------------------------------------------------------------------------------
    ''''    Convert captured (serial) bit stream to data waveform, Assume LSB->MSB in the bit stream (reversed
    ''''        order may be easily accommodated by adding a switch in the argument list)
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        BitWf2Arry is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        InWf          : DSP Wave (serial) to be converted
    ''''        WrdWdth  : number of bits per word
    ''''        NoOfSamples    : number of samples found in the bit stream
    ''''        DataWf         : converted (parallel) DSP Wave
    ''''

''    NoOfSamples = InWf.SampleSize

''    If NoOfSamples Mod WrdWdth <> 0 Then
''         Debug.Print vbNewLine & "Bit stream wave size not integer times of the word width." _
''            & " Waveform will Be truncated" & vbNewLine
''    End If

    DataWf_Binary = InWf.Select(StartIndex, , WrdWdth).COPY
    DataWf = InWf.Select(StartIndex, , WrdWdth).ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

End Function
Public Function Split_Dspwave(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
        For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
End Function

Public Function CreatSerDigSrc(ByVal DataWf As DSPWave, ByVal BitWdthWf As DSPWave, ByVal SrcSize As Long, _
                        ByVal isIndDataRepeat As Boolean, ByVal isAllDataRepeat As Boolean, _
                        ByRef DigSrcWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    This function generates DSP waveform for DSSC digital source, to replace Create_DigSrc_Data,
    ''''        which was written in VBT module, not DSP. Other changes included remove all hard codings,
    ''''        increase the flexibility for various inputs and repetition combinations. The number of
    ''''        data (registers) are now flexible
    ''''
    ''''    This function is for serial source only. If parallel, the method of generating the waveform
    ''''        should be different for best efficiency
    ''''
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/25/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''        DataWf As DSPWave       Data to be sourced.  Each element (paralle, long) for each register
    ''''        BitWdthWf As DSPWave    Bit width info, should have the same number of element as in DataWf
    ''''        SrcSize As Long         If serial: total number of bits. If parallel, number of data
    ''''        isIndDataRepeat/isAllDataRepeat As Boolean  repeating each data/reg or repeat complete sequence
    ''''                                one must be true while the other false
    ''''        DigSrcWf As DSPWave     The generated waveform to be used for digital source
    
    Dim NoOfData As Long, NoOfRepeats As Long
    Dim DataIdx As Long
    Dim tmpwf As New DSPWave
    Dim SingleWordWf As New DSPWave         '''' hold one single data point
    
    '''' re-initiate the dig source wf
    Set DigSrcWf = New DSPWave
    
    ''''    Verifiy the number of samples in the data wf and bitwidth wf are the same
    NoOfData = DataWf.SampleSize
    If NoOfData <> BitWdthWf.SampleSize Then
        Debug.Print vbNewLine & "The sizes of data wf and width wf no matching!" & vbNewLine
    End If
    
    ''''    Calculate the number of repeats and check if integer
    NoOfRepeats = SrcSize \ BitWdthWf.CalcSum
    If NoOfRepeats * BitWdthWf.CalcSum <> SrcSize Then
        Debug.Print vbNewLine & "The total number of source bits not integer of the sum of the data widths!" & vbNewLine
    End If
    
    ''''
    '''' creat waveform based on the repeat scheme: per data repeat or repeat all
    '''' must be done one data at a time because the bitwdiths of the data may be different
    ''''
    If isIndDataRepeat Then
        For DataIdx = 0 To NoOfData - 1
            SingleWordWf = DataWf.Select(DataIdx, 1, 1).ConvertStreamTo( _
                tldspSerial, BitWdthWf.Element(DataIdx), 0, Bit0IsLsb)     '''' one data at a time. 1: stride. 1: size
            
            '''' concatenation only works on wafevorm with more than one element
            If DigSrcWf.SampleSize > 0 Then
                DigSrcWf = DigSrcWf.Concatenate(SingleWordWf.Repeat(NoOfRepeats))
            Else
                DigSrcWf = SingleWordWf.Repeat(NoOfRepeats)
            End If
        Next DataIdx
    ElseIf isAllDataRepeat Then
        For DataIdx = 0 To NoOfData - 1
            SingleWordWf = DataWf.Select(DataIdx, 1, 1).ConvertStreamTo( _
                tldspSerial, BitWdthWf.Element(DataIdx), 0, Bit0IsLsb)     '''' one data at a time. 1: stride. 1: size
            
            If DigSrcWf.SampleSize > 0 Then
                DigSrcWf = DigSrcWf.Concatenate(SingleWordWf)
            Else
                DigSrcWf = SingleWordWf.COPY
            End If
        Next DataIdx
        DigSrcWf = DigSrcWf.Repeat(NoOfRepeats)
    End If
    
    
End Function
Public Function CreateFlexibleDSPWave(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long
    
    ''Dim SerialStream As New DSPWave
    ''Dim ParallelStream As New DSPWave
    ''
    ''ParallelStream.CreateConstant 17, 1, DspLong
    ''SerialStream = ParallelStream.ConvertStreamTo(tldspSerial, 12, 0, Bit0IsLsb)
    InWf = InWf.ConvertDataTypeTo(DspLong)
    DataWf = InWf.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsLsb)
    
End Function
Public Function CreateFlexibleDSPWave_lpro(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave, ByRef fine_wf As DSPWave) As Long
    
    ''Dim SerialStream As New DSPWave
    ''Dim ParallelStream As New DSPWave
    ''
    ''ParallelStream.CreateConstant 17, 1, DspLong
    ''SerialStream = ParallelStream.ConvertStreamTo(tldspSerial, 12, 0, Bit0IsLsb)
    InWf = InWf.ConvertDataTypeTo(DspLong)
    DataWf = InWf.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    DataWf = fine_wf.Concatenate(DataWf)
    
End Function

Public Function SetupTrimCodeBit(ByVal InWf As DSPWave, ByVal b_SetupToBit0 As Boolean, ByVal BitIndex As Long, ByVal b_ControlNextBit As Boolean, ByRef DataWf As DSPWave) As Long
    
    DataWf = InWf.COPY
    If b_SetupToBit0 = True Then
        DataWf.Element(BitIndex) = 0
    Else
        DataWf.Element(BitIndex) = 1
    End If
    If b_ControlNextBit Then
        DataWf.Element(BitIndex - 1) = 1
    End If
End Function

Public Function SetupTrimCodeBit_rev(ByVal InWf As DSPWave, ByVal b_SetupToBit0 As Boolean, ByVal BitIndex As Long, ByVal b_ControlNextBit As Boolean, ByRef DataWf As DSPWave) As Long
    
    DataWf = InWf
    If b_SetupToBit0 = True Then
        DataWf.Element(BitIndex) = 1
    Else
        DataWf.Element(BitIndex) = 0
    End If
    If b_ControlNextBit Then
        DataWf.Element(BitIndex - 1) = 1
    End If
End Function

Public Function ConvertToLongAndSerialToParrel(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long
    
    InWf = InWf.ConvertDataTypeTo(DspLong)
    
     If (InWf.SampleSize > 1) Then 'Check for Decimal...If binary convert to decimal or else copies the decimal directly to output waveform
        DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
    Else
        DataWf = InWf.COPY
    End If
   ' DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
End Function
Public Function CombineDSPWave_and_ConvertToLongAndSerialToParrel(ByVal FirstDSP As DSPWave, ByVal SecondDSP As DSPWave, _
                                                                  ByVal FirstLength As Long, ByVal SecondLength As Long, ByRef CombineDSP As DSPWave, _
                                                                  ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long
    

    CombineDSP.CreateConstant 0, FirstLength + SecondLength, DspLong
    FirstDSP = FirstDSP.ConvertDataTypeTo(DspLong)
    SecondDSP = SecondDSP.ConvertDataTypeTo(DspLong)
    CombineDSP.Select(0, 1, FirstLength).Replace (FirstDSP)
    CombineDSP.Select(FirstLength, 1, SecondLength).Replace (SecondDSP)
    
    CombineDSP = CombineDSP.ConvertDataTypeTo(DspLong)
    DataWf = CombineDSP.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    
End Function
Public Function FindMaxEyeWidth_reverse(ByVal SwqEyeWf As DSPWave, ByVal SwsEyeWf As DSPWave, _
                        ByVal NoOfBists As Integer, ByRef EyeWidthWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Search for the eyewidth corresponding to the largest eye in the presense of one more eyes.
    ''''    An eye is defined as a continuous "1's"
    ''''    Based on the test methodologies used for TMA, there are 2 sweeps from the center of the UI, the first
    ''''        towards left, and second right, Swq and Sws respectively. The function would stitch them together by
    ''''        reversing the Swq waveform, and concatenating with the Sws wave.
    ''''
    ''''    The resulting eye diagram is from the left to right covering the entire UI.

    
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016

    ''''    Usage
    ''''        FindMaxEyeWidth is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        SwqEyeWf,SwsEyeWf   : Eye diagram waveforms, all BISTs concatenated.
    ''''        NoOfBists           : DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        EyeWidthWf          : Array of the eyewidths found, in DSP wave format, to be passed back
    
    ''''    rev 1, by Tim, 2/26/2016
    ''''
    ''''--------------------------------------------------------------------------------------------------
    ''''           New add to fix out of index error, exit for not an good method. Need to change it.
    ''''--------------------------------------------------------------------------------------------------

    Dim WholeEyeWf As New DSPWave
    Dim NoOfSamples As Long, EyeStrobes As Long
    Dim EyeWidth As Long
    Dim BistIdx As Long, idx As Long, idx2 As Long
    Dim MaxEye As Long
    EyeWidthWf.CreateConstant 0, NoOfBists, DspLong
    
    NoOfSamples = SwqEyeWf.SampleSize
    If NoOfSamples <> SwsEyeWf.SampleSize Then
             Debug.Print vbNewLine & "The lengths of Swq and Sws eye sweep waveform not consistent!" & vbNewLine
    End If
    
    EyeStrobes = NoOfSamples / NoOfBists
    
    For BistIdx = 0 To NoOfBists - 1
        WholeEyeWf = SwqEyeWf.Select(BistIdx * EyeStrobes, 1, EyeStrobes).COPY   '''' dummy operation, to allocate element
        
        '''' SWQ is from right to left, starting the middle. Reversing the eye diagram
        For idx = 0 To EyeStrobes - 1
            idx2 = (BistIdx + 1) * EyeStrobes - 1 - idx
            WholeEyeWf.Element(idx) = SwqEyeWf.Element(idx2)
        Next idx
        
        '''' stitch the reversed SWQ eye diagram to the SWK one
        WholeEyeWf = WholeEyeWf.Concatenate(SwsEyeWf.Select(BistIdx * EyeStrobes, 1, EyeStrobes).COPY)
        
        MaxEye = 0
        Dim EyeResultFlag As Boolean
        
        
        
        '''' finding the maximum eyewidth
        For idx = 0 To 2 * EyeStrobes - 1
            EyeResultFlag = False
            If WholeEyeWf.Element(idx) = 1 Then '''' starting a new eye
                EyeWidth = 0
                Do
                    EyeWidth = EyeWidth + 1
                    idx = idx + 1
               
                If idx >= 2 * EyeStrobes Then
                    EyeResultFlag = True
                ElseIf WholeEyeWf.Element(idx) = 0 Then
                    EyeResultFlag = True
                Else 'Do nothing '20230601
                End If
               
               
               'Loop 'Until WholeEyeWf.Element(idx) = 0
                        
              '  If WholeEyeWf.Element(idx) = 0 Then EyeResultFlag = True
                
                
                'Loop Until idx >= 2 * EyeStrobes - 1 Or WholeEyeWf.Element(idx) = 0
                Loop Until EyeResultFlag = True
                
                    '''' >= workaround for not being able to escape the loop while maintain the wave index within the range
                    ''''    in the condition
                    ''' The next statement separately including the last strobe
                If idx = 2 * EyeStrobes - 1 Then EyeWidth = EyeWidth + WholeEyeWf.Element(idx)
                If MaxEye < EyeWidth Then MaxEye = EyeWidth
            End If
        Next idx
        
        EyeWidthWf.Element(BistIdx) = MaxEye
        
    Next BistIdx
    

End Function

Public Function DSP_Add(ByRef InWf_1 As DSPWave, ByRef InWf_2 As DSPWave) As Long
    
    InWf_1 = InWf_1.Add(InWf_2)
    
End Function

Public Function DSP_Subtract(ByRef InWf_1 As DSPWave, ByRef InWf_2 As DSPWave) As Long
    
    InWf_1 = InWf_1.Subtract(InWf_2)
    
End Function

Public Function DSP_Multiply(ByRef InWf_1 As DSPWave, ByRef InWf_2 As DSPWave) As Long
    
    InWf_1 = InWf_1.Multiply(InWf_2)
    
End Function

Public Function DSP_Divide(ByRef InWf_1 As DSPWave, ByRef InWf_2 As DSPWave) As Long
    
    InWf_1 = InWf_1.divide(InWf_2)
    
End Function

Public Function DSPWaveDecToBinary(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long
    InWf = InWf.ConvertDataTypeTo(DspLong)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    DataWf = InWf.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsMsb)
    
End Function

Public Function Transfer2GrayCode(ByVal InWf As DSPWave, ByRef OutWf As DSPWave, ByRef OutWf_DEC As DSPWave) As Long
''    Exit Function
    OutWf.CreateConstant 0, InWf.SampleSize, DspLong
    OutWf_DEC.CreateConstant 0, 1, DspLong
    
    Dim i As Long
    For i = 0 To InWf.SampleSize - 1 Step 1
        If i = 0 Then
            OutWf.Element(i) = InWf.Element(i)
        Else
            If InWf.Element(i - 1) = InWf.Element(i) Then
                OutWf.Element(i) = 0
            Else
                OutWf.Element(i) = 1
            End If
        End If
    Next i
    OutWf_DEC = OutWf.ConvertStreamTo(tldspParallel, OutWf.SampleSize, 0, Bit0IsMsb)
    
End Function

Public Function PreCheckMinMaxTrimCode(ByVal b_SetupToBit0 As Boolean, ByRef DataWf As DSPWave) As Long
    
    Dim i As Long
    For i = 0 To DataWf.SampleSize - 1
        If b_SetupToBit0 = True Then
            DataWf.Element(i) = 0
        Else
            DataWf.Element(i) = 1
        End If
    Next i
End Function

Public Function DSP_DivideConstant(ByRef InWf_1 As DSPWave, ByVal Denominator As Long) As Long
    If Denominator = 0 Then
''        TheExec.Datalog.WriteComment ("Error! Divide 0.")
        Exit Function
    Else
        InWf_1 = InWf_1.divide(Denominator)
    End If
End Function


Public Function CombineDSPWave(ByVal FirstDSP As DSPWave, ByVal SecondDSP As DSPWave, ByVal FirstLength As Long, ByVal SecondLength As Long, ByRef CombineDSP As DSPWave) As Long
    Dim i As Long, j As Long
    Dim index As Long
    CombineDSP.CreateConstant 0, FirstLength + SecondLength, DspLong
    FirstDSP = FirstDSP.ConvertDataTypeTo(DspLong)
    SecondDSP = SecondDSP.ConvertDataTypeTo(DspLong)
    CombineDSP.Select(0, 1, FirstLength).Replace (FirstDSP)
    CombineDSP.Select(FirstLength, 1, SecondLength).Replace (SecondDSP)
End Function

Public Function SelectCertainBitsToDec(ByVal InWf As DSPWave, ByVal StartBit As Long, ByVal BitLength As Long, ByRef DataWf As DSPWave) As Long
    Dim TempDSP As New DSPWave
    InWf = InWf.ConvertDataTypeTo(DspLong)
    TempDSP = InWf.Select(StartBit, 1, BitLength).COPY
    
    DataWf = TempDSP.ConvertStreamTo(tldspParallel, BitLength, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
End Function
Public Function ConcatenateDSP(ByVal DSPWave_First As DSPWave, ByVal First_StartElement As Long, ByVal First_EndElement As Long, _
                               ByVal DSPWave_Second As DSPWave, ByVal Second_StartElement As Long, ByVal Second_EndElement As Long, _
                               ByRef DSPWave_Combine As DSPWave) As Long

    Dim FinalLength As Long
    Dim i As Long
    FinalLength = Abs(First_EndElement - First_StartElement) + Abs(Second_EndElement - Second_StartElement) + 2
    DSPWave_Combine.CreateConstant 0, FinalLength
    
    Dim b_MinToMax_First As Boolean
    Dim b_MinToMax_Second As Boolean
    Dim Step_First As Integer
    Dim Step_Second As Integer
    Dim counter As Long
    counter = 0
    If First_EndElement - First_StartElement > 0 Then
        b_MinToMax_First = True
        Step_First = 1
    Else
        b_MinToMax_First = False
        Step_First = -1
    End If
    
    If Second_EndElement - Second_StartElement > 0 Then
        b_MinToMax_Second = True
        Step_Second = 1
    Else
        b_MinToMax_Second = False
        Step_Second = -1
    End If
    
    For i = First_StartElement To First_EndElement Step Step_First
        DSPWave_Combine.Element(counter) = DSPWave_First.Element(i)
        counter = counter + 1
    Next i


    For i = Second_StartElement To Second_EndElement Step Step_Second
        DSPWave_Combine.Element(counter) = DSPWave_Second.Element(i)
        counter = counter + 1
    Next i
End Function

Public Function DSP_BitWiseAnd(ByVal InputDSP As DSPWave, ByVal FixedDSP As DSPWave, ByVal bitwidth As Long, ByRef OutputDSP As DSPWave) As Long
    InputDSP = InputDSP.ConvertDataTypeTo(DspLong)
    OutputDSP.CreateConstant 0, bitwidth, DspLong
    OutputDSP = InputDSP.bitwiseand(FixedDSP)
End Function


Public Function DSP_BitWiseOr(ByVal InputDSP As DSPWave, ByVal FixedDSP As DSPWave, ByVal bitwidth As Long, ByRef OutputDSP As DSPWave) As Long
    InputDSP = InputDSP.ConvertDataTypeTo(DspLong)
    OutputDSP.CreateConstant 0, bitwidth, DspLong
    OutputDSP = InputDSP.BitwiseOr(FixedDSP)
End Function

Public Function DSP_BitWiseXOR(ByVal InputDSP As DSPWave, ByVal FixedDSP As DSPWave, ByVal bitwidth As Long, ByRef OutputDSP As DSPWave) As Long
    InputDSP = InputDSP.ConvertDataTypeTo(DspLong)
    OutputDSP.CreateConstant 0, bitwidth, DspLong
    OutputDSP = InputDSP.BitwiseXor(FixedDSP)
End Function


Public Function BinToDec(ByVal InWf As DSPWave, ByRef DataWf As DSPWave) As Long
    Dim WrdWdth As Long
    InWf = InWf.ConvertDataTypeTo(DspLong)
    WrdWdth = InWf.SampleSize
    DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

End Function

Public Function DSP_Convert_2S_Complement(ByVal InWf As DSPWave, WordWidth As Long, DataWf As DSPWave) As Long
    InWf = InWf.ConvertDataTypeTo(DspLong)
    InWf.WordWidth = WordWidth
    Set DataWf = InWf.ConvertDataTypeTo(DspLong)
''    Debug.Print DataWf.Element(0)
    
End Function

Sub WordWidthExample()
Dim FourBitValues As New DSPWave
Dim IntegerValues As New DSPWave
Dim i As Long
FourBitValues.CreateRamp 0, 1, 16, DspLong
FourBitValues.WordWidth = 4
Set IntegerValues = FourBitValues.ConvertDataTypeTo(DspLong)
For i = 0 To 15
    Debug.Print FourBitValues.Element(i); " => "; IntegerValues.Element(i)
Next
End Sub
Public Function DSP_ConvertDataTypeToLong(ByRef InWf_1 As DSPWave) As Long
    InWf_1.ConvertDataTypeTo (DspLong)
End Function

Public Function DSP_GrayCode2Bin(ByVal IsUnsigned As Boolean, ByVal InWf As DSPWave, ByRef OutWf As DSPWave, ByRef OutWf_DEC As DSPWave) As Long

    OutWf.CreateConstant 0, InWf.SampleSize, DspLong
    OutWf_DEC.CreateConstant 0, 1, DspLong
    
    Dim i As Long
    Dim MSB_ElementNumForSignUnsign As Long
    MSB_ElementNumForSignUnsign = InWf.SampleSize - 1
    
    Dim SignUnsignDiffBit As Long
    If IsUnsigned Then
        SignUnsignDiffBit = 1
    Else
        SignUnsignDiffBit = 2
    End If
    Dim index As Long
    index = 0
    For i = InWf.SampleSize - SignUnsignDiffBit To 0 Step -1
        If index = 0 Then
            OutWf.Element(i) = InWf.Element(i)
        Else
            If InWf.Element(i) = OutWf.Element(i + 1) Then
                OutWf.Element(i) = 0
            Else
                OutWf.Element(i) = 1
            End If
        End If
        index = index + 1
    Next i
    OutWf_DEC = OutWf.ConvertStreamTo(tldspParallel, OutWf.SampleSize, 0, Bit0IsMsb)
    
    If IsUnsigned = True Then
    
    Else
        If InWf.Element(MSB_ElementNumForSignUnsign) = 1 Then
            OutWf_DEC.Element(0) = OutWf_DEC.Element(0) * -1
        Else
        End If
    End If
    
End Function

Public Function DSP_2S_Complement_To_SignDec(InWf As DSPWave, WordWidth As Long, DataWf_DEC As DSPWave) As Long
''    Exit Function
    
    InWf = InWf.ConvertDataTypeTo(DspLong)
''''    InWf.WordWidth = WordWidth
''    Set DataWf = InWf.ConvertNumFormatTo(SignMagnitude, WordWidth)
''    Debug.Print DataWf.Element(0)
    Dim MaxVal As Double
    MaxVal = 2 ^ (WordWidth - 1)
    Dim DSPWaveWithoutSignBit As New DSPWave
    Dim SignBit As Long
    DSPWaveWithoutSignBit = InWf.Select(0, , WordWidth - 1).COPY
    SignBit = InWf.Element(WordWidth - 1)
    Dim DSPWaveWithoutSignBit_DEC As New DSPWave
    DSPWaveWithoutSignBit_DEC.CreateConstant 0, 1, DspLong
''    Call rundsp.BinToDec(DSPWaveWithoutSignBit, DSPWaveWithoutSignBit_DEC)
    Dim WrdWdthWithoutSignBit As Long
    ''InWf = InWf.ConvertDataTypeTo(DspLong)
    WrdWdthWithoutSignBit = DSPWaveWithoutSignBit.SampleSize
    DSPWaveWithoutSignBit_DEC = DSPWaveWithoutSignBit.ConvertStreamTo(tldspParallel, WrdWdthWithoutSignBit, 0, Bit0IsMsb)
    
    If SignBit = 0 Then
        DataWf_DEC = DSPWaveWithoutSignBit_DEC.COPY
    Else
        DataWf_DEC = DSPWaveWithoutSignBit_DEC.Subtract(MaxVal)
    End If
End Function

Public Function DSP_SignedBin_To_SignDec(InWf As DSPWave, WordWidth As Long, DataWf_DEC As DSPWave) As Long
''    Exit Function
    
    InWf = InWf.ConvertDataTypeTo(DspLong)
''''    InWf.WordWidth = WordWidth
''    Set DataWf = InWf.ConvertNumFormatTo(SignMagnitude, WordWidth)
''    Debug.Print DataWf.Element(0)
    Dim MaxVal As Long
    MaxVal = 2 ^ (WordWidth - 1)
    Dim DSPWaveWithoutSignBit As New DSPWave
    Dim SignBit As Long
    'DSPWaveWithoutSignBit = InWf.Select(0, , WordWidth - 1).Copy
    DSPWaveWithoutSignBit = InWf.Select(0, , WordWidth).COPY
    SignBit = InWf.Element(WordWidth - 1)
    Dim DSPWaveWithoutSignBit_DEC As New DSPWave
    DSPWaveWithoutSignBit_DEC.CreateConstant 0, 1, DspLong
''    Call rundsp.BinToDec(DSPWaveWithoutSignBit, DSPWaveWithoutSignBit_DEC)
    Dim WrdWdthWithoutSignBit As Long
    ''InWf = InWf.ConvertDataTypeTo(DspLong)
    WrdWdthWithoutSignBit = DSPWaveWithoutSignBit.SampleSize
    DSPWaveWithoutSignBit_DEC = DSPWaveWithoutSignBit.ConvertStreamTo(tldspParallel, WrdWdthWithoutSignBit, 0, Bit0IsMsb)
    If MaxVal > DSPWaveWithoutSignBit_DEC.Element(0) Then
        DataWf_DEC = DSPWaveWithoutSignBit_DEC
    Else
        'DataWf_DEC = DSPWaveWithoutSignBit_DEC.Negate
        DataWf_DEC = DSPWaveWithoutSignBit_DEC.Negate.Add(MaxVal)
        'DataWf_DEC = MaxVal.Subtract(DSPWaveWithoutSignBit_DEC).Negate
    End If
End Function

Public Function LPDPRX_EyeSweep(ByVal InputDSPWave As DSPWave, ByVal FinalEyeOutBitNum As Long, ByRef CalcOutputDSPWave As DSPWave, ByRef CalcEyeWidth As Long) As Long
    Dim i As Long
    Dim index As Long
    Dim MaxWitth As Long
    Dim TempMaxWidth As Long
    Dim FinalMaxWitth As Long
    
    CalcOutputDSPWave.CreateConstant 0, FinalEyeOutBitNum, DspLong
    
    For i = 0 To InputDSPWave.SampleSize - 1 Step 2
        If InputDSPWave.Element(i) = 0 And InputDSPWave.Element(i + 1) = 32768 Then
            CalcOutputDSPWave.Element(index) = 0
        Else
            CalcOutputDSPWave.Element(index) = 1
        End If
        index = index + 1
    Next i
'    CalcOutputDSPWave.Element(1) = 0
'    CalcOutputDSPWave.Element(2) = 0
'    CalcOutputDSPWave.Element(10) = 0
'    CalcOutputDSPWave.Element(11) = 0
'    CalcOutputDSPWave.Element(12) = 0
'    CalcOutputDSPWave.Element(21) = 0
'    CalcOutputDSPWave.Element(22) = 0
    MaxWitth = 0
    Dim All_zero_Flag As Boolean
    All_zero_Flag = True
    
    For i = 0 To CalcOutputDSPWave.SampleSize - 1
        If CalcOutputDSPWave.Element(i) = 0 Then
            MaxWitth = MaxWitth + 1
            TempMaxWidth = MaxWitth
        Else
            All_zero_Flag = False
            If FinalMaxWitth < TempMaxWidth Then
                FinalMaxWitth = TempMaxWidth
            End If
            MaxWitth = 0
        End If
    Next i
    If All_zero_Flag = True Then
        FinalMaxWitth = TempMaxWidth
    End If
    
    CalcEyeWidth = FinalMaxWitth
End Function

Public Function PCIE_EyeSweep(ByVal InputDSPWave As DSPWave, ByVal FinalEyeOutBitNum As Long, ByRef CalcOutputDSPWave As DSPWave, ByRef CalcEyeWidth As Long) As Long
    Dim i As Long
    Dim index As Long
    Dim MaxWitth As Long
    Dim TempMaxWidth As Long
    Dim FinalMaxWitth As Long
    
    CalcOutputDSPWave.CreateConstant 0, FinalEyeOutBitNum, DspLong
    
    For i = 0 To InputDSPWave.SampleSize - 1 Step 1
        If InputDSPWave.Element(i) = 0 Then 'And InputDSPWave.Element(i + 1) = 0 Then
            CalcOutputDSPWave.Element(index) = 0
        Else
            CalcOutputDSPWave.Element(index) = 1
        End If
        index = index + 1
    Next i
'    CalcOutputDSPWave.Element(1) = 0
'    CalcOutputDSPWave.Element(2) = 0
'    CalcOutputDSPWave.Element(10) = 0
'    CalcOutputDSPWave.Element(11) = 0
'    CalcOutputDSPWave.Element(12) = 0
'    CalcOutputDSPWave.Element(21) = 0
'    CalcOutputDSPWave.Element(22) = 0
    MaxWitth = 0
    Dim All_zero_Flag As Boolean
    All_zero_Flag = True
    
    For i = 0 To CalcOutputDSPWave.SampleSize - 1
        If CalcOutputDSPWave.Element(i) = 0 Then
            MaxWitth = MaxWitth + 1
            TempMaxWidth = MaxWitth
        Else
            All_zero_Flag = False
            If FinalMaxWitth < TempMaxWidth Then
                FinalMaxWitth = TempMaxWidth
            End If
            MaxWitth = 0
        End If
    Next i
    If All_zero_Flag = True Then
        FinalMaxWitth = TempMaxWidth
    End If
    
    CalcEyeWidth = FinalMaxWitth
End Function


Public Function SeprateDSP(DSP_Input As DSPWave, DSP_Input_UpperBIN As DSPWave, DSP_Input_BelowBIN As DSPWave) As Long
    Dim length As Long
    length = DSP_Input.SampleSize / 2
    DSP_Input_UpperBIN = DSP_Input.Select(0, , length).COPY
    DSP_Input_BelowBIN = DSP_Input.Select(0 + length, , length).COPY
End Function

Public Function AveWithStdev(InputWave As DSPWave, mean As Double, Std As Double) As Long
    mean = InputWave.CalcMeanWithStdDev(Std)
End Function


Public Function SeprateDSP_TTR_Single(DSP_Input_Update As DSPWave, DSP_Input_UpperBIN As DSPWave, DSP_Input_BelowBIN As DSPWave, ByRef DSP_Input_UpperDEC As DSPWave, ByRef DSP_Input_BelowDEC As DSPWave) As Long
 
    Dim length As Long
    Dim i As Long
    Dim WrdWdth As Long
    length = 8
    
    For i = 0 To DSP_Input_Update.SampleSize - 1 Step 16
        If i = 0 Then
            DSP_Input_UpperBIN = DSP_Input_Update.Select(i, , length).COPY
            DSP_Input_BelowBIN = DSP_Input_Update.Select(i + length, , length).COPY
        Else
            DSP_Input_UpperBIN = DSP_Input_UpperBIN.Concatenate(DSP_Input_Update.Select(i, , length).COPY)
            DSP_Input_BelowBIN = DSP_Input_BelowBIN.Concatenate(DSP_Input_Update.Select(i + length, , length).COPY)
        End If
    Next i
            
    DSP_Input_UpperDEC = DSP_Input_UpperBIN.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_BelowDEC = DSP_Input_BelowBIN.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    
End Function
Public Function SeprateDSP_TTR(DSP_Input_Update As DSPWave, DSP_Input_UpperBIN_1 As DSPWave, DSP_Input_UpperBIN_2 As DSPWave, DSP_Input_BelowBIN_1 As DSPWave, DSP_Input_BelowBIN_2 As DSPWave, ByRef DSP_Input_UpperDEC_1 As DSPWave, ByRef DSP_Input_UpperDEC_2 As DSPWave, ByRef DSP_Input_BelowDEC_1 As DSPWave, ByRef DSP_Input_BelowDEC_2 As DSPWave) As Long
 
    Dim length As Long
    Dim i As Long
    Dim WrdWdth As Long
    length = 8

    For i = 0 To DSP_Input_Update.SampleSize - 1 Step 32
        If i = 0 Then
            DSP_Input_UpperBIN_1 = DSP_Input_Update.Select(i, , length).COPY
            DSP_Input_UpperBIN_2 = DSP_Input_Update.Select(i + length, , length).COPY
            DSP_Input_BelowBIN_1 = DSP_Input_Update.Select(i + length * 2, , length).COPY
            DSP_Input_BelowBIN_2 = DSP_Input_Update.Select(i + length * 3, , length).COPY
        Else
            DSP_Input_UpperBIN_1 = DSP_Input_UpperBIN_1.Concatenate(DSP_Input_Update.Select(i, , length).COPY)
            DSP_Input_UpperBIN_2 = DSP_Input_UpperBIN_2.Concatenate(DSP_Input_Update.Select(i + length, , length).COPY)
            DSP_Input_BelowBIN_1 = DSP_Input_BelowBIN_1.Concatenate(DSP_Input_Update.Select(i + length * 2, , length).COPY)
            DSP_Input_BelowBIN_2 = DSP_Input_BelowBIN_2.Concatenate(DSP_Input_Update.Select(i + length * 3, , length).COPY)
        
        End If
    Next i
            
    DSP_Input_UpperDEC_1 = DSP_Input_UpperBIN_1.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_UpperDEC_2 = DSP_Input_UpperBIN_2.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_BelowDEC_1 = DSP_Input_BelowBIN_1.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_BelowDEC_2 = DSP_Input_BelowBIN_2.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    
End Function
Public Function FindMaxEyeWidth_reverse_bywidth(ByVal SwqEyeWf As DSPWave, ByVal SwsEyeWf As DSPWave, _
                        ByVal Cont_width As DSPWave, ByRef EyeWidthWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Search for the eyewidth corresponding to the largest eye in the presense of one more eyes.
    ''''    An eye is defined as a continuous "1's"
    ''''    Based on the test methodologies used for TMA, there are 2 sweeps from the center of the UI, the first
    ''''        towards left, and second right, Swq and Sws respectively. The function would stitch them together by
    ''''        reversing the Swq waveform, and concatenating with the Sws wave.
    ''''
    ''''    The resulting eye diagram is from the left to right covering the entire UI.

    
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016

    ''''    Usage
    ''''        FindMaxEyeWidth is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        SwqEyeWf,SwsEyeWf   : Eye diagram waveforms, all BISTs concatenated.
    ''''        NoOfBists           : DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        EyeWidthWf          : Array of the eyewidths found, in DSP wave format, to be passed back
    
    ''''    rev 1, by Tim, 2/26/2016
    ''''
    ''''--------------------------------------------------------------------------------------------------
    ''''           New add to fix out of index error, exit for not an good method. Need to change it.
    ''''--------------------------------------------------------------------------------------------------

    Dim WholeEyeWf As New DSPWave
    Dim NoOfSamples As Long, EyeStrobes As Long
    Dim EyeWidth As Long
    Dim BistIdx As Long, idx As Long, idx2 As Long
    Dim MaxEye As Long
    Dim lane_count As Long
    Dim EyeStrobes_split As Long
    Dim Count_Eyerunbit As Long
    
    lane_count = Cont_width.SampleSize
    
    
    NoOfSamples = SwqEyeWf.SampleSize
    
    If NoOfSamples <> SwsEyeWf.SampleSize Then
             Debug.Print vbNewLine & "The lengths of Swq and Sws eye sweep waveform not consistent!" & vbNewLine
        
    End If
    
   EyeWidthWf.CreateConstant 0, lane_count, DspLong
    
    For BistIdx = 0 To lane_count - 1
    
       EyeStrobes = Cont_width.Element(BistIdx)
       
            If BistIdx = 0 Then
                
                  EyeStrobes_split = 0
                  
                  Count_Eyerunbit = Cont_width.Element(BistIdx)
                
            Else
                
                  EyeStrobes_split = EyeStrobes_split + Cont_width.Element(BistIdx - 1)
                  
                  Count_Eyerunbit = Count_Eyerunbit + Cont_width.Element(BistIdx)
'                  Count_Eyerunbit = Count_Eyerunbit + Cont_width.Element(BistIdx - 1)
            End If
        
        WholeEyeWf = SwqEyeWf.Select(EyeStrobes_split, 1, EyeStrobes).COPY    '''' dummy operation, to allocate element
        
        '''' SWQ is from right to left, starting the middle. Reversing the eye diagram
        For idx = 0 To EyeStrobes - 1
            idx2 = Count_Eyerunbit - 1 - idx
            WholeEyeWf.Element(idx) = SwqEyeWf.Element(idx2)
        Next idx
        
        '''' stitch the reversed SWQ eye diagram to the SWK one
        WholeEyeWf = WholeEyeWf.Concatenate(SwsEyeWf.Select(EyeStrobes_split, 1, EyeStrobes).COPY)
        
        MaxEye = 0
        Dim EyeResultFlag As Boolean
        
        
        
        '''' finding the maximum eyewidth
        For idx = 0 To 2 * EyeStrobes - 1
            EyeResultFlag = False
            If WholeEyeWf.Element(idx) = 1 Then '''' starting a new eye
                EyeWidth = 0
                Do
                    EyeWidth = EyeWidth + 1
                    idx = idx + 1
               
                If idx >= 2 * EyeStrobes Then
                    EyeResultFlag = True
                ElseIf WholeEyeWf.Element(idx) = 0 Then
                    EyeResultFlag = True
                Else 'Do nothing '20230601
                End If
               
               
               'Loop 'Until WholeEyeWf.Element(idx) = 0
                        
              '  If WholeEyeWf.Element(idx) = 0 Then EyeResultFlag = True
                
                
                'Loop Until idx >= 2 * EyeStrobes - 1 Or WholeEyeWf.Element(idx) = 0
                Loop Until EyeResultFlag = True
                
                    '''' >= workaround for not being able to escape the loop while maintain the wave index within the range
                    ''''    in the condition
                    ''' The next statement separately including the last strobe
                If idx = 2 * EyeStrobes - 1 Then EyeWidth = EyeWidth + WholeEyeWf.Element(idx)
                If MaxEye < EyeWidth Then MaxEye = EyeWidth
            End If
        Next idx
        
        EyeWidthWf.Element(BistIdx) = MaxEye
        
    Next BistIdx
End Function

Public Function DspWaveMergeRepeat(ByRef OutputDspWave As DSPWave, ByVal InDSPwave As DSPWave, ByVal SampleSize As Long) As Long
    
    Dim InDspWaveSampleSize As Long
    
    InDspWaveSampleSize = InDSPwave.SampleSize
    OutputDspWave = InDSPwave.Repeat(CLng(SampleSize / InDspWaveSampleSize))
End Function

Public Function DSPWf_Concatenate(OutputDspWave As DSPWave, InDSPwave As DSPWave, dummy As Long) As Long

    If OutputDspWave.SampleSize = 0 Then
        OutputDspWave = InDSPwave.COPY
    Else
        OutputDspWave = OutputDspWave.ConvertDataTypeTo(DspLong)
        InDSPwave = InDSPwave.ConvertDataTypeTo(DspLong)
        OutputDspWave = OutputDspWave.Concatenate(InDSPwave)
    End If

End Function

Public Function MTR_ASGMTR_Freq_Calculation(ByRef InWf1 As DSPWave, ByRef InWf2 As DSPWave, ByRef InWf3 As DSPWave, ByRef InWf4 As DSPWave, ByRef InWf5 As DSPWave, ByRef InWf6 As DSPWave, ByRef InWf7 As DSPWave, ByRef InWf8 As DSPWave, ByRef InWf9 As DSPWave, ByRef InWf10 As DSPWave, ByRef InWf11 As DSPWave, ByRef InWf12 As DSPWave, ByRef InWf13 As DSPWave, ByRef InWf14 As DSPWave, ByRef InWf15 As DSPWave, ByRef InWf16 As DSPWave, ByRef InWf17 As DSPWave, ByRef InWf18 As DSPWave, ByRef InWf19 As DSPWave, ByRef InWf20 As DSPWave, ByRef InWf21 As DSPWave, ByRef InWf22 As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long

Dim DataWf1 As New DSPWave: DataWf1 = InWf1.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf2 As New DSPWave: DataWf2 = InWf2.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf3 As New DSPWave: DataWf3 = InWf3.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf4 As New DSPWave: DataWf4 = InWf4.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf5 As New DSPWave: DataWf5 = InWf5.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf6 As New DSPWave: DataWf6 = InWf6.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf7 As New DSPWave: DataWf7 = InWf7.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf8 As New DSPWave: DataWf8 = InWf8.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf9 As New DSPWave: DataWf9 = InWf9.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf10 As New DSPWave: DataWf10 = InWf10.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf11 As New DSPWave: DataWf11 = InWf11.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf12 As New DSPWave: DataWf12 = InWf12.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf13 As New DSPWave: DataWf13 = InWf13.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf14 As New DSPWave: DataWf14 = InWf14.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf15 As New DSPWave: DataWf15 = InWf15.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf16 As New DSPWave: DataWf16 = InWf16.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf17 As New DSPWave: DataWf17 = InWf17.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf18 As New DSPWave: DataWf18 = InWf18.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf19 As New DSPWave: DataWf19 = InWf19.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf20 As New DSPWave: DataWf20 = InWf20.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf21 As New DSPWave: DataWf21 = InWf21.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf22 As New DSPWave: DataWf22 = InWf22.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)

InWf1 = InWf1.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf2 = InWf2.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf3 = InWf3.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf4 = InWf4.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf5 = InWf5.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf6 = InWf6.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf7 = InWf7.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf8 = InWf8.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf9 = InWf9.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf10 = InWf10.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf11 = InWf11.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf12 = InWf12.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf13 = InWf13.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf14 = InWf14.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf15 = InWf15.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf16 = InWf16.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf17 = InWf17.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf18 = InWf18.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf19 = InWf19.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf20 = InWf20.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf21 = InWf21.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf22 = InWf22.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

DataWf = DataWf1.Concatenate(DataWf2).Concatenate(DataWf3).Concatenate(DataWf4).Concatenate(DataWf5).Concatenate(DataWf6).Concatenate(DataWf7).Concatenate(DataWf8).Concatenate(DataWf9).Concatenate(DataWf10).Concatenate(DataWf11).Concatenate(DataWf12).Concatenate(DataWf13).Concatenate(DataWf14).Concatenate(DataWf15).Concatenate(DataWf16).Concatenate(DataWf17).Concatenate(DataWf18).Concatenate(DataWf19).Concatenate(DataWf20).Concatenate(DataWf21).Concatenate(DataWf22)


End Function
Public Function Split_Dspwave_PCIETXPLL_Central_bak(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, Delta_Value As DSPWave, target_var As Double, ByVal calibration_target_value As Long, ByVal start_search As Long, BinTarget As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    '=====================================================20180904
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    
    ''''calc and print in datalog
'    temp1_dict.CreateConstant 0, 1, DspLong
'    temp2_dict.CreateConstant 0, 1, DspLong
'
'    For i = 0 To CLng((OutWf.SampleSize) / 2 - 1)
'            temp1_dict.Element(0) = OutWf.Element(i)
'            temp2_dict.Element(0) = OutWf.Element(i + 32)
'            calc_data.Element(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
'    Next i
    j = 1
    For i = start_search To start_search + CLng(OutWf.SampleSize) - 2
        calc_data.Element(i - start_search) = OutWf.Element(j)
        j = j + 1
    Next i
    ''' compare the target

    temp_delta_value = 9999
    j = 1
    For k = start_search To start_search + CLng((OutWf.SampleSize)) - 2
        Delta_Value.Element(j) = Abs(calibration_target_value - calc_data.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(j) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(j)
            target_var = k
        End If
        j = j + 1
    Next k
    
    ''''''''''''' Decimal to Binary dspwave for dictionary'''''''''''''''
    
    Dim TempVal As Long
    
    TempVal = target_var
    For i = 0 To CLng((BinTarget.SampleSize)) - 1
        BinTarget.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            BinTarget.Element(i) = 0
        End If
    Next i

        
    '=====================================================20180904
    
    
End Function

Public Function Split_Dspwave_PCIETXPLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, Delta_Value As DSPWave, target_var As Double, ByVal calibration_target_value As Long, ByVal start_search As Long, BinTarget As DSPWave, delta_value2 As DSPWave, target_var2 As Double, BinTarget2 As DSPWave, ByVal FirstTrimPoint As Integer) As Long
    '20211227 CW modify for Ibiza
    
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    '=====================================================20180904
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    'calc_data.CreateConstant 0, OutWf.SampleSize - 1, DspDouble
    
    For i = start_search To CLng(OutWf.SampleSize) - 1
        calc_data.Element(i - start_search) = OutWf.Element(i)
    Next i

    temp_delta_value = 9999
                 
    For k = start_search To CLng(OutWf.SampleSize) - 1 Step 2
        Delta_Value.Element(k) = Abs(calibration_target_value - calc_data.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var = k / 2
        End If
    Next k
    
    temp_delta_value = 9999
                 
    For k = (start_search + 1) To CLng(OutWf.SampleSize) - 1 Step 2
        delta_value2.Element(k) = Abs(calibration_target_value - calc_data.Element(k - start_search))
        'search min delta
        If delta_value2.Element(k) < temp_delta_value Then
            temp_delta_value = delta_value2.Element(k)
            target_var2 = (k - 1) / 2
        End If
    Next k
    
    target_var = target_var + FirstTrimPoint '200425 by CW, first trim point for TTR
    target_var2 = target_var2 + FirstTrimPoint '200425 by CW, first trim point for TTR
    
    ''''''''''''' Decimal to Binary dspwave for dictionary'''''''''''''''
    
    Dim TempVal As Long
    
    TempVal = target_var
    For i = 0 To CLng((BinTarget.SampleSize)) - 1
        BinTarget.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            BinTarget.Element(i) = 0
        End If
    Next i
    
    TempVal = target_var2
    For i = 0 To CLng((BinTarget2.SampleSize)) - 1
        BinTarget2.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            BinTarget2.Element(i) = 0
        End If
    Next i

        
    '=====================================================20180904
    
    
End Function

Public Function MTR_DSGMTR_Freq_Calculation(ByRef InWf1 As DSPWave, ByRef InWf2 As DSPWave, ByRef InWf3 As DSPWave, ByRef InWf4 As DSPWave, ByRef InWf5 As DSPWave, ByRef InWf6 As DSPWave, ByRef InWf7 As DSPWave, ByRef InWf8 As DSPWave, ByRef InWf9 As DSPWave, ByRef InWf10 As DSPWave, ByRef InWf11 As DSPWave, ByRef InWf12 As DSPWave, ByRef InWf13 As DSPWave, ByRef InWf14 As DSPWave, ByRef InWf15 As DSPWave, ByRef InWf16 As DSPWave, ByRef InWf17 As DSPWave, ByRef InWf18 As DSPWave, ByRef InWf19 As DSPWave, ByRef InWf20 As DSPWave, ByRef InWf21 As DSPWave, ByRef InWf22 As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long

Dim DataWf1 As New DSPWave: DataWf1 = InWf1.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf2 As New DSPWave: DataWf2 = InWf2.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf3 As New DSPWave: DataWf3 = InWf3.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf4 As New DSPWave: DataWf4 = InWf4.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf5 As New DSPWave: DataWf5 = InWf5.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf6 As New DSPWave: DataWf6 = InWf6.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf7 As New DSPWave: DataWf7 = InWf7.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf8 As New DSPWave: DataWf8 = InWf8.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf9 As New DSPWave: DataWf9 = InWf9.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf10 As New DSPWave: DataWf10 = InWf10.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf11 As New DSPWave: DataWf11 = InWf11.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf12 As New DSPWave: DataWf12 = InWf12.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf13 As New DSPWave: DataWf13 = InWf13.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf14 As New DSPWave: DataWf14 = InWf14.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf15 As New DSPWave: DataWf15 = InWf15.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf16 As New DSPWave: DataWf16 = InWf16.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf17 As New DSPWave: DataWf17 = InWf17.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf18 As New DSPWave: DataWf18 = InWf18.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf19 As New DSPWave: DataWf19 = InWf19.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf20 As New DSPWave: DataWf20 = InWf20.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf21 As New DSPWave: DataWf21 = InWf21.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf22 As New DSPWave: DataWf22 = InWf22.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)

InWf1 = InWf1.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf2 = InWf2.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf3 = InWf3.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf4 = InWf4.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf5 = InWf5.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf6 = InWf6.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf7 = InWf7.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf8 = InWf8.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf9 = InWf9.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf10 = InWf10.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf11 = InWf11.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf12 = InWf12.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf13 = InWf13.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf14 = InWf14.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf15 = InWf15.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf16 = InWf16.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf17 = InWf17.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf18 = InWf18.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf19 = InWf19.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf20 = InWf20.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf21 = InWf21.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf22 = InWf22.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

DataWf = DataWf1.Concatenate(DataWf2).Concatenate(DataWf3).Concatenate(DataWf4).Concatenate(DataWf5).Concatenate(DataWf6).Concatenate(DataWf7).Concatenate(DataWf8).Concatenate(DataWf9).Concatenate(DataWf10).Concatenate(DataWf11).Concatenate(DataWf12).Concatenate(DataWf13).Concatenate(DataWf14).Concatenate(DataWf15).Concatenate(DataWf16).Concatenate(DataWf17).Concatenate(DataWf18).Concatenate(DataWf19).Concatenate(DataWf20).Concatenate(DataWf21).Concatenate(DataWf22)


End Function

Public Function Split_Dspwave_CIOCALC(Outwf_T1 As DSPWave, v2 As Double, V3 As Double, V4 As Double, V5 As Double, V6 As Double, V7 As Double, V8 As Double, V9 As Double, V10 As Double, V11 As Double, V12 As Double, V13 As Double, V14 As Double, V15 As Double, V16 As Double, V17 As Double, V18 As Double, V19 As Double, V20 As Double, V21 As Double, V22 As Double, V23 As Double, _
                                V24 As Double, V25 As Double, storeDSP As DSPWave) As Long
                                                                                                                                                                                                                                                               
                                                                                                                                                                                                                                                               
    ' Special calculation for T1
    ' 1.1 store measured data
    Dim d_temp(24) As Double
    d_temp(0) = v2
    d_temp(1) = V3
    d_temp(2) = V4
    d_temp(3) = V5
    d_temp(4) = V6
    d_temp(5) = V7
    d_temp(6) = V8
    d_temp(7) = V9
    d_temp(8) = V10
    d_temp(9) = V11
    d_temp(10) = V12
    d_temp(11) = V13
    d_temp(12) = V14
    d_temp(13) = V15
    d_temp(14) = V16
    d_temp(15) = V17
    d_temp(16) = V18
    d_temp(17) = V19
    d_temp(18) = V20
    d_temp(19) = V21
    d_temp(20) = V22
    d_temp(21) = V23
    d_temp(22) = V24
    d_temp(23) = V25
    ' 1.2 Calculate desired index
    Dim target_index As Long: target_index = 999
    Dim target_gap As Double: target_gap = 2013144
    Dim i As Long
    For i = 0 To 23
        If d_temp(i) >= 0.3 And (Outwf_T1.Element(i) - 3333) >= 0 Then
            target_gap = Abs(Outwf_T1.Element(i) - 3333)
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next
    
    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_gap = 2013144
        For i = 0 To 23
            If d_temp(i) >= 0.3 And Abs(Outwf_T1.Element(i) - 3333) < target_gap Then
                target_gap = Abs(Outwf_T1.Element(i) - 3333)
                target_index = i
            End If
        Next
    End If
'
    Outwf_T1.Element(24) = target_index
    
    Dim TempVal As Long
    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T1.Element(24) = TempVal
    
    For i = 0 To CLng((storeDSP.SampleSize)) - 1
        storeDSP.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            storeDSP.Element(i) = 0
        End If
    Next i
                                                                                                                                                                                                                                                               
End Function

Public Function Split_Dspwave_CIOPLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, Delta_Value As DSPWave, target_var_low As Double, target_var_high As Double, target_var_low2 As Double, target_var_high2 As Double, ByVal calibration_target_value_low As Long, ByVal calibration_target_value_high As Long, ByVal start_search As Long, Bin_Target_low As DSPWave, Bin_Target_high As DSPWave, Bin_Target_low2 As DSPWave, Bin_Target_high2 As DSPWave) As Long
     'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    '=====================================================20180904
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    
    ''''calc and print in datalog
'    temp1_dict.CreateConstant 0, 1, DspLong
'    temp2_dict.CreateConstant 0, 1, DspLong
'
'    For i = 0 To CLng((OutWf.SampleSize) / 2 - 1)
'            temp1_dict.Element(0) = OutWf.Element(i)
'            temp2_dict.Element(0) = OutWf.Element(i + 32)
'            calc_data.Element(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
'    Next i
    For i = 0 To CLng(OutWf.SampleSize) - 1
        calc_data.Element(i - start_search) = OutWf.Element(i)
    Next i
    ''' compare the target

    temp_delta_value = 9999
    
    For k = start_search To CLng((OutWf.SampleSize)) - 1 Step 2
        Delta_Value.Element(k) = Abs(calibration_target_value_low - OutWf.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var_low = k / 2
        End If
    Next k
    
    
    temp_delta_value = 9999

    For k = start_search To CLng((OutWf.SampleSize)) - 1 Step 2
        Delta_Value.Element(k) = Abs(calibration_target_value_high - OutWf.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var_high = k / 2
        End If
    Next k
    

    temp_delta_value = 9999
    
    For k = (start_search + 1) To CLng((OutWf.SampleSize)) - 1 Step 2
        Delta_Value.Element(k) = Abs(calibration_target_value_low - OutWf.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var_low2 = (k + 1) / 2
        End If
    Next k

    
    temp_delta_value = 9999

    For k = (start_search + 1) To CLng((OutWf.SampleSize)) - 1 Step 2
        Delta_Value.Element(k) = Abs(calibration_target_value_high - OutWf.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var_high2 = (k + 1) / 2
        End If
    Next k
    
    
'    '==========Skua 1 trim==========
'
'    For k = start_search To CLng((OutWf.SampleSize)) - 1
'        delta_value.Element(k) = Abs(calibration_target_value_low - OutWf.Element(k))
'        'search min delta
'        If delta_value.Element(k) < temp_delta_value Then
'            temp_delta_value = delta_value.Element(k)
'            target_var_low = k
'        End If
'    Next k
'
'    temp_delta_value = 9999
'
'    For k = start_search To CLng((OutWf.SampleSize)) - 1
'        delta_value.Element(k) = Abs(calibration_target_value_high - OutWf.Element(k))
'        'search min delta
'        If delta_value.Element(k) < temp_delta_value Then
'            temp_delta_value = delta_value.Element(k)
'            target_var_high = k
'        End If
'    Next k
'
'    '==========Skua 1 trim==========
    
    '=====================================================20180904
    Dim TempVal As Long
    '@220102 TTR updated by Walker
'    target_var_low = target_var_low + 26
'    target_var_low2 = target_var_low2 + 26
'    target_var_high = target_var_high + 26
'    target_var_high2 = target_var_high2 + 26
    
    TempVal = target_var_low
    For i = 0 To CLng((Bin_Target_low.SampleSize)) - 1
        Bin_Target_low.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_low.Element(i) = 0
        End If
    Next i

    TempVal = target_var_high
    For i = 0 To CLng((Bin_Target_low.SampleSize)) - 1
        Bin_Target_high.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_high.Element(i) = 0
        End If
    Next i
    
    TempVal = target_var_low2
    For i = 0 To CLng((Bin_Target_low2.SampleSize)) - 1
        Bin_Target_low2.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_low2.Element(i) = 0
        End If
    Next i

    TempVal = target_var_high2
    For i = 0 To CLng((Bin_Target_low2.SampleSize)) - 1
        Bin_Target_high2.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_high2.Element(i) = 0
        End If
    Next i
    
End Function


Public Function Split_Dspwave_AUS(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, Outwf_T1 As DSPWave, sda_measuredata3 As Double, sda_measuredata4 As Double, sda_measuredata5 As Double, sda_measuredata6 As Double, sda_measuredata7 As Double, sda_measuredata8 As Double, sda_measuredata9 As Double, sda_measuredata10 As Double, sda_measuredata11 As Double, sda_measuredata12 As Double, sda_measuredata13 As Double, sda_measuredata14 As Double, sda_measuredata15 As Double, sda_measuredata16 As Double, sda_measuredata17 As Double, sda_measuredata18 As Double, sda_measuredata19 As Double, sda_measuredata20 As Double, sda_measuredata21 As Double, sda_measuredata22 As Double, sda_measuredata23 As Double, sda_measuredata24 As Double, sda_measuredata25 As Double, _
                                sda_measuredata26 As Double, storeDSP As DSPWave) As Long
                                                                                                                                                                                                                                                               
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim temp_dsp As New DSPWave
    Dim Count As Long
    current_loc = 0
    Count = 0
                                                                                                                                                                                                                                                                '' 20170705 - Solve for over 32 bits to decimal
    temp_dsp.CreateConstant 0, width_Wf.SampleSize, DspDouble
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        If width_Wf.Element(i) = 16 Then
            'dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
            temp_dsp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
            'current_loc = current_loc + width_Wf.Element(i)
        End If
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    Outwf_T1 = temp_dsp.Select(1, 1, 25).COPY
                                                                                                                                                                                                                                                               
    ' Special calculation for T1
    ' 1.1 store measured data
    Dim d_temp(24) As Double
    d_temp(0) = sda_measuredata3
    d_temp(1) = sda_measuredata4
    d_temp(2) = sda_measuredata5
    d_temp(3) = sda_measuredata6
    d_temp(4) = sda_measuredata7
    d_temp(5) = sda_measuredata8
    d_temp(6) = sda_measuredata9
    d_temp(7) = sda_measuredata10
    d_temp(8) = sda_measuredata11
    d_temp(9) = sda_measuredata12
    d_temp(10) = sda_measuredata13
    d_temp(11) = sda_measuredata14
    d_temp(12) = sda_measuredata15
    d_temp(13) = sda_measuredata16
    d_temp(14) = sda_measuredata17
    d_temp(15) = sda_measuredata18
    d_temp(16) = sda_measuredata19
    d_temp(17) = sda_measuredata20
    d_temp(18) = sda_measuredata21
    d_temp(19) = sda_measuredata22
    d_temp(20) = sda_measuredata23
    d_temp(21) = sda_measuredata24
    d_temp(22) = sda_measuredata25
    d_temp(23) = sda_measuredata26
    ' 1.2 Calculate desired index
    Dim target_index As Long: target_index = 999
    Dim target_gap As Double: target_gap = 2013144
    For i = 0 To 23
        If d_temp(i) >= 0.3 And (Outwf_T1.Element(i) - 3333) >= 0 Then
            target_gap = Abs(Outwf_T1.Element(i) - 3333)
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next
    
    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_gap = 2013144
        For i = 0 To 23
            If d_temp(i) >= 0.3 And Abs(Outwf_T1.Element(i) - 3333) < target_gap Then
                target_gap = Abs(Outwf_T1.Element(i) - 3333)
                target_index = i
            End If
        Next
    End If
'
    Outwf_T1.Element(24) = target_index
    
    Dim TempVal As Long
    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T1.Element(24) = TempVal
    
    For i = 0 To CLng((storeDSP.SampleSize)) - 1
        storeDSP.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            storeDSP.Element(i) = 0
        End If
    Next i
                                                                                                                                                                                                                                                               
End Function

Public Function Split_Dspwave_CIO(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, Outwf_T1 As DSPWave, Outwf_T2 As DSPWave) As Long
    
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim temp_dsp As New DSPWave
    Dim Count As Long
    current_loc = 0
    Count = 0
    
    temp_dsp.CreateConstant 0, width_Wf.SampleSize, DspDouble
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        If width_Wf.Element(i) = 16 Then
            temp_dsp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
        End If

        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    Outwf_T1 = temp_dsp.Select(1, 1, Count / 2 - 2).COPY
    Outwf_T2 = temp_dsp.Select(Count / 2 + 1, 1, Count / 2 - 2).COPY
    
End Function

Public Function Split_Dspwave_PCIEREFPLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, ByRef Outwf_T1 As DSPWave, ByRef sda_measuredata As DSPWave, calc_data As DSPWave, Delta_Value As DSPWave, target_var As Double, ByVal calibration_target_value As Long, ByVal start_search As Long, ByVal search_width As Long, BinTarget As DSPWave, _
                ByVal Interspersed_boolen As Boolean, ByVal z As Long, ByVal Lane_Sum As Long, ByVal offset As Long) As Long

    'Dim Split_Wave_ary(2) As New DSPWave
    
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim Interspersed_DSP As New DSPWave
    Dim Outwf_T1_Temp As New DSPWave
    Dim Count As Long
    Count = 0
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    Outwf_T1_Temp.CreateConstant 0, sda_measuredata.Element(0) * Lane_Sum, DspLong
    For i = 0 To width_Wf.SampleSize - 1
            dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
            If width_Wf.Element(i) = search_width Then
                Outwf_T1_Temp.Element(Count) = dec_wave.Element(0)
                Count = Count + 1
            End If
            OutWf.Element(i) = dec_wave.Element(0)
            current_loc = current_loc + width_Wf.Element(i)
    Next i
     
    If Interspersed_boolen = True Then
       Interspersed_DSP = Outwf_T1_Temp.Select(z, Lane_Sum, sda_measuredata.Element(0)).COPY
      Else
       Interspersed_DSP = Outwf_T1_Temp.Select(z * sda_measuredata.Element(0), 1, sda_measuredata.Element(0)).COPY
    End If
    
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    
    ''''calc and print in datalog
'    temp1_dict.CreateConstant 0, 1, DspLong
'    temp2_dict.CreateConstant 0, 1, DspLong
'
'    For i = 0 To CLng((OutWf.SampleSize) / 2 - 1)
'            temp1_dict.Element(0) = OutWf.Element(i)
'            temp2_dict.Element(0) = OutWf.Element(i + 32)
'            calc_data.Element(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
'    Next i
    For i = start_search To CLng(Interspersed_DSP.SampleSize) - 1
        calc_data.Element(i - start_search) = Interspersed_DSP.Element(i)
    Next i
    ''' compare the target

    temp_delta_value = 9999
                 
    For k = start_search To CLng((Interspersed_DSP.SampleSize)) - 1
        Delta_Value.Element(k) = Abs(calibration_target_value - calc_data.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var = k
        End If
    Next k
    target_var = target_var + offset
    
    ''''''''''''' Decimal to Binary dspwave for dictionary'''''''''''''''
    
    Dim TempVal As Long
    
    TempVal = target_var
    For i = 0 To CLng((BinTarget.SampleSize)) - 1
        BinTarget.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            BinTarget.Element(i) = 0
        End If
    Next i

        
    '=====================================================20180904
    
    
End Function
Public Function Split_Dspwave_CIOREFPLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, Delta_Value As DSPWave, target_var As Double, ByVal calibration_target_value As Long, ByVal start_search As Long, BinTarget As DSPWave, ByVal FirstTrimPoint As Integer) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    '=====================================================20180904
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    calc_data.CreateConstant 0, OutWf.SampleSize - 1, DspDouble '
    
    ''''calc and print in datalog
'    temp1_dict.CreateConstant 0, 1, DspLong
'    temp2_dict.CreateConstant 0, 1, DspLong
'
'    For i = 0 To CLng((OutWf.SampleSize) / 2 - 1)
'            temp1_dict.Element(0) = OutWf.Element(i)
'            temp2_dict.Element(0) = OutWf.Element(i + 32)
'            calc_data.Element(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
'    Next i
    For i = start_search To CLng(OutWf.SampleSize) - 1
        calc_data.Element(i - start_search) = OutWf.Element(i)
    Next i
    ''' compare the target

    temp_delta_value = 9999
                 
    For k = start_search To CLng((OutWf.SampleSize)) - 1
        Delta_Value.Element(k) = Abs(calibration_target_value - calc_data.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var = k
        End If
    Next k
    
    '@220102 TTR updated by Walker
'    target_var = target_var + FirstTrimPoint '200425 by CW, first trim point for TTR
    
    ''''''''''''' Decimal to Binary dspwave for dictionary'''''''''''''''
    
    Dim TempVal As Long
    
    TempVal = target_var
    For i = 0 To CLng((BinTarget.SampleSize)) - 1
        BinTarget.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            BinTarget.Element(i) = 0
        End If
    Next i

        
    '=====================================================20180904
    
    
End Function

Public Function Split_2SComplementDSPWave_To_SignDec(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim DSP_2sComplement As New DSPWave
    current_loc = 0
    For i = 0 To width_Wf.SampleSize - 1
        DSP_2sComplement = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        Call DSP_2S_Complement_To_SignDec(DSP_2sComplement, width_Wf.Element(i), dec_wave)
        If i = 0 Then
            OutWf = dec_wave.COPY
        Else
            OutWf = OutWf.Concatenate(dec_wave)
        End If
        current_loc = current_loc + width_Wf.Element(i)
    Next i
End Function

Public Function Split_Gray_2sComplementDSPWave_to_Dec(DSPSignedGray_StartBit As DSPWave, DSPUnSignedGray_StartBit As DSPWave, DSP2sComplement_StartBit As DSPWave, DSPSignedBin_StartBit As DSPWave, InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
Dim dec_wave As New DSPWave, current_loc As Long, i As Long
Dim Index_SignedGray As Long: Index_SignedGray = 0
Dim Index_UnSignedGray As Long: Index_UnSignedGray = 0
Dim Index_2sComplement As Long: Index_2sComplement = 0
Dim Index_SignedBin As Long: Index_SignedBin = 0
Dim DSP_SignedGray As New DSPWave
Dim DSP_UnSignedGray As New DSPWave
Dim DSP_2sComplement As New DSPWave
Dim DSP_SignedBin As New DSPWave
Dim DSP_UnSignedBinary As New DSPWave
Dim Out_Wf_Binary As New DSPWave
current_loc = 0
For i = 0 To width_Wf.SampleSize - 1
    If current_loc = DSPSignedGray_StartBit.Element(Index_SignedGray) Then
        DSP_SignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        Call DSP_GrayCode2Bin(False, DSP_SignedGray, Out_Wf_Binary, dec_wave)
        If Index_SignedGray <> DSPSignedGray_StartBit.SampleSize - 1 Then
            Index_SignedGray = Index_SignedGray + 1
        End If
    ElseIf current_loc = DSPUnSignedGray_StartBit.Element(Index_UnSignedGray) Then
        DSP_UnSignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        Call DSP_GrayCode2Bin(True, DSP_UnSignedGray, Out_Wf_Binary, dec_wave)
        If Index_UnSignedGray <> DSPUnSignedGray_StartBit.SampleSize - 1 Then
            Index_UnSignedGray = Index_UnSignedGray + 1
        End If
    ElseIf current_loc = DSP2sComplement_StartBit.Element(Index_2sComplement) Then
        DSP_2sComplement = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        Call DSP_2S_Complement_To_SignDec(DSP_2sComplement, width_Wf.Element(i), dec_wave)
        If Index_2sComplement <> DSP2sComplement_StartBit.SampleSize - 1 Then
            Index_2sComplement = Index_2sComplement + 1
        End If
    ElseIf current_loc = DSPSignedBin_StartBit.Element(Index_SignedBin) Then
        DSP_SignedBin = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        Call DSP_SignedBin_To_SignDec(DSP_SignedBin, width_Wf.Element(i), dec_wave)
        If Index_SignedBin <> DSPSignedBin_StartBit.SampleSize - 1 Then
            Index_SignedBin = Index_SignedBin + 1
        End If
    Else
        DSP_UnSignedBinary = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        dec_wave = DSP_UnSignedBinary.ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
    End If
    If i = 0 Then
        OutWf = dec_wave.COPY
    Else
        OutWf = OutWf.Concatenate(dec_wave)
    End If
    current_loc = current_loc + width_Wf.Element(i)
Next i
End Function

Public Function DSP_Opt_EYE(ByRef DqSwpWf As DSPWave, ByVal DSP_Eye_StartBit_DQ As DSPWave, ByVal DSP_Eye_BitLength_DQ As DSPWave, ByRef DqsSwpWf As DSPWave, ByVal DSP_Eye_StartBit_DQS As DSPWave, ByVal DSP_Eye_BitLength_DQS As DSPWave, ByVal NoOfBists As Integer, ByRef DQ_EYE_Data As DSPWave, ByRef DQS_EYE_Data As DSPWave, ByRef DSP_Eye_Width As DSPWave) As Long

    Dim i As Integer

    For i = 0 To DSP_Eye_StartBit_DQ.SampleSize - 1
        If i = 0 Then
            DQ_EYE_Data = DqSwpWf.Select(DSP_Eye_StartBit_DQ.Element(i), , DSP_Eye_BitLength_DQ.Element(i)).COPY
        Else
             DQ_EYE_Data = DQ_EYE_Data.Concatenate(DqSwpWf.Select(DSP_Eye_StartBit_DQ.Element(i), , DSP_Eye_BitLength_DQ.Element(i)).COPY)
        End If
    Next i

    For i = 0 To DSP_Eye_StartBit_DQS.SampleSize - 1
        If i = 0 Then
            DQS_EYE_Data = DqsSwpWf.Select(DSP_Eye_StartBit_DQS.Element(i), , DSP_Eye_BitLength_DQS.Element(i)).COPY
        Else
            DQS_EYE_Data = DQS_EYE_Data.Concatenate(DqsSwpWf.Select(DSP_Eye_StartBit_DQS.Element(i), , DSP_Eye_BitLength_DQS.Element(i)).COPY)
        End If
    Next i

    Call FindMaxEyeWidth_reverse(DQ_EYE_Data, DQS_EYE_Data, NoOfBists, DSP_Eye_Width)
    
End Function

Public Function Split_Gray_to_Dec(DSPSignedGray_StartBit As DSPWave, DSPUnSignedGray_StartBit As DSPWave, InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
Dim dec_wave As New DSPWave, current_loc As Long, i As Long
Dim Index_SignedGray As Long: Index_SignedGray = 0
Dim Index_UnSignedGray As Long: Index_UnSignedGray = 0
Dim Index_2sComplement As Long: Index_2sComplement = 0
Dim DSP_SignedGray As New DSPWave
Dim DSP_UnSignedGray As New DSPWave
Dim DSP_2sComplement As New DSPWave
Dim DSP_UnSignedBinary As New DSPWave
Dim Out_Wf_Binary As New DSPWave
current_loc = 0
For i = 0 To width_Wf.SampleSize - 1
    If current_loc = DSPSignedGray_StartBit.Element(Index_SignedGray) Then
        DSP_SignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        Call DSP_GrayCode2Bin(False, DSP_SignedGray, Out_Wf_Binary, dec_wave)
        If Index_SignedGray <> DSPSignedGray_StartBit.SampleSize - 1 Then
            Index_SignedGray = Index_SignedGray + 1
        End If
    ElseIf current_loc = DSPUnSignedGray_StartBit.Element(Index_UnSignedGray) Then
        DSP_UnSignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        Call DSP_GrayCode2Bin(True, DSP_UnSignedGray, Out_Wf_Binary, dec_wave)
        If Index_UnSignedGray <> DSPUnSignedGray_StartBit.SampleSize - 1 Then
            Index_UnSignedGray = Index_UnSignedGray + 1
        End If
'    ElseIf current_loc = DSP2sComplement_StartBit.Element(Index_2sComplement) Then
'        DSP_2sComplement = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
'        Call DSP_2S_Complement_To_SignDec(DSP_2sComplement, width_Wf.Element(i), dec_wave)
'        If Index_2sComplement <> DSP2sComplement_StartBit.SampleSize - 1 Then
'            Index_2sComplement = Index_2sComplement + 1
'        End If
    Else
        DSP_UnSignedBinary = InWf.Select(current_loc, , width_Wf.Element(i)).COPY
        dec_wave = DSP_UnSignedBinary.ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
    End If
    If i = 0 Then
        OutWf = dec_wave.COPY
    Else
        OutWf = OutWf.Concatenate(dec_wave)
    End If
    current_loc = current_loc + width_Wf.Element(i)
Next i
End Function

''20190604AddFunction
Public Function DSPWf_Dec2Binary(ByVal InWf As DSPWave, ByVal DataWdth As Long, ByRef OutWf As DSPWave) As Long
    
    OutWf = InWf.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, DataWdth, 0, Bit0IsMsb)
    
End Function
Public Function ReAssignmentDSPWave(ByVal InWf As DSPWave, ByVal SeparateNum As Long, ByRef OutWf As DSPWave, ByVal Src_dig As Boolean, ByVal assignment As DSPWave, ByVal AssignmentDSPWave As DSPWave) As Long
    Dim i As Integer    'Dylan Edited 20190528
    Dim j As Integer
    Dim RegSize As Long
    Dim InWfSize As Long
    Dim FullSize As Long
    
    
    InWfSize = InWf.SampleSize
    FullSize = OutWf.SampleSize
    RegSize = FullSize / SeparateNum
    
    If Src_dig = False Then
        For i = 1 To SeparateNum
            For j = 1 To RegSize
                OutWf.Element((i * RegSize) - j) = InWf.Element(RegSize - j)
                'Debug.Print i * RegSize - j
            Next j
        Next i
    Else
        For i = 1 To SeparateNum
            If AssignmentDSPWave.Element(i - 1) = 0 Then
               For j = 1 To RegSize
                OutWf.Element((i * RegSize) - j) = InWf.Element(RegSize - j)
                'Debug.Print i * RegSize - j
               Next j
            Else
               For j = 1 To RegSize
               OutWf.Element((i * RegSize) - j) = assignment.Element(RegSize - j)
               Next j
            End If
        Next
      
    End If
      
End Function

Public Function ElementTransformer(ByRef InWf As DSPWave, ByVal SeparateNum As Long, ByVal ElementOffset As Long) As Long
    Dim i As Long       'Dylan Edited 20190528
    Dim SizeCnt As Long
    Dim ProcessDSP As New DSPWave
    SizeCnt = InWf.SampleSize
    ProcessDSP.CreateConstant 0, SizeCnt, DspLong
    
    If ElementOffset = 0 Then
        ElementOffset = ElementOffset - 1
        For i = 0 To SizeCnt - 1
            ProcessDSP.Element(i) = InWf.Element(ElementOffset - i)
        Next i
        
        For i = 0 To SizeCnt - 1
            InWf.Element(ElementOffset - SeparateNum + i) = ProcessDSP.Element(i)
        Next i
    Else
        For i = 0 To SizeCnt - 1
            ProcessDSP.Element(i) = InWf.Element(SizeCnt - i - 1)
        Next i
        
        For i = 0 To SizeCnt - 1
            InWf.Element(i) = ProcessDSP.Element(i)
        Next i
    End If
    
    
    
    
    
    
    
'''''    j = SeparateNum / 2
'''''    SeparateNum = SeparateNum - 1                           ' Minimum element is zero
'''''    ElementOffset = ElementOffset - 1                       ' MAxmum element is Maxmum - 1
'''''    ProcessDSP.CreateConstant 0, SeparateNum, DspLong
'''''
'''''
'''''    For i = 0 To j - 1
'''''        ProcessDSP.Element(i) = InWf.Element(ElementOffset - i)
'''''        InWf.Element(ElementOffset - i) = InWf.Element(ElementOffset - SeparateNum + i)
'''''        InWf.Element(ElementOffset - SeparateNum + i) = ProcessDSP.Element(i)
'''''    Next i

End Function
Public Function SetupLinearTrimCodeBit(ByVal TrimMethod As Boolean, ByRef TrimCode As Double, ByVal b_SetupToBit0 As Boolean, _
ByVal RegSize As Long, ByRef DataWf As DSPWave, ByVal doallFlag As Boolean) As Long
' Dylan Edited 20190529
    Dim i As Integer
    Dim SizeTemp As Integer
    Dim TrimTemp As Integer
    Dim TotallySize As Integer
    
    SizeTemp = CInt(RegSize) - 1
    TotallySize = CInt(DataWf.SampleSize) - 1
    
    If TrimMethod = True Then                                   ' Increase Linear
        If doallFlag = True Then
            TrimCode = TrimCode + 1
        ElseIf b_SetupToBit0 = False Then
            TrimCode = TrimCode + 1
        End If
    Else                                                        ' Decrease Linear
        If doallFlag = True Then
            TrimCode = TrimCode - 1
        ElseIf b_SetupToBit0 = False Then
            TrimCode = TrimCode - 1
        End If
    End If
    
    TrimTemp = TrimCode
    If b_SetupToBit0 = False Or doallFlag = True Then
        For i = 0 To SizeTemp
            If TrimTemp <> 0 Then
                If TrimTemp \ (2 ^ (SizeTemp - i)) <> 0 Then
                    DataWf.Element(TotallySize - i) = 1
                Else
                    DataWf.Element(TotallySize - i) = 0
                End If
                If i <> SizeTemp Then
                    If TrimTemp >= (2 ^ (SizeTemp - i)) Then
                        TrimTemp = TrimTemp - (2 ^ (SizeTemp - i))
                    End If
                End If
            Else
                DataWf.Element(TotallySize - i) = 0
            End If
        Next i
    End If
    
End Function

Public Function CalculateDSPWaveforTrimCode(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave, _
ByVal TrimTotallyOffset As Integer, ByVal TrimBasedNum As Integer, ByRef FinallyWf As DSPWave) As Long
' Dylan Edited 20190529
    Dim i As Integer
    Dim CalculateDSPWave As New DSPWave
    
    If InWf.SampleSize <> 1 Then                                                                    ' Avoid sweep fail which any site
        CalculateDSPWave.CreateConstant 0, 1, DspLong
        InWf = InWf.ConvertDataTypeTo(DspLong)
        DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
        DataWf.Element(0) = TrimBasedNum - (TrimTotallyOffset - DataWf.Element(0))
        
        
        If DataWf.Element(0) >= 0 Then                                                              ' Transfer to Binary if element(0) is positive
            CalculateDSPWave.Element(0) = DataWf.Element(0)
            FinallyWf = CalculateDSPWave.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsMsb)        ' Transfer from Decimal to Binary
        Else                                                                                        ' Transfer to 2'S if element(0) is negative
            CalculateDSPWave.Element(0) = Abs(DataWf.Element(0))
            If CalculateDSPWave.Element(0) > (2 ^ (WrdWdth - 1) - 1) Then                           ' Illegal jugment
                FinallyWf.CreateConstant 0, 1, DspLong
            Else
                CalculateDSPWave.Element(0) = DataWf.Element(0)
                CalculateDSPWave.Element(0) = CalculateDSPWave.Element(0) + 2 ^ WrdWdth             ' Calculate 2'S complement
                FinallyWf = CalculateDSPWave.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsMsb)    ' Transfer from Binary to Decimal
            End If
        End If
    End If
End Function

Public Function SetupBinaryTrimCodeBit(ByVal InWf As DSPWave, ByVal b_SetupToBit0 As Boolean, ByVal BitIndex As Long, ByVal InitStateByCapCode As Long, _
ByVal TrimOffset As Long, ByVal TrimOriginalSize As Long, ByRef DataWf As DSPWave, ByVal b_ControlNextBit As Boolean, ByVal AssignmentDSPWave As DSPWave, _
ByVal CoverSize As Long) As Long
'Dylan Edited 20190615
       
    Dim i As Long
    
    Dim Square As Long
    Dim SelsetOffset As Long
    Dim CalculateSize As Long
    Dim CalculateDSP As New DSPWave
    Dim assigment_select As Long
    CalculateSize = CLng(DataWf.SampleSize)
    DataWf = InWf.COPY

    For i = 0 To AssignmentDSPWave.SampleSize
        If AssignmentDSPWave.Element(i) = 0 Then
            SelsetOffset = i * CalculateSize
            Exit For
        Else
            SelsetOffset = 0
        End If
    Next i


    CalculateDSP = DataWf.Select(SelsetOffset, 1, CalculateSize).COPY
    DataWf = CalculateDSP.COPY
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    DataWf = DataWf.ConvertStreamTo(tldspParallel, CalculateSize, 0, Bit0IsMsb)
    If DataWf.Element(0) <> 0 Then
        DataWf.Element(0) = DataWf.Element(0) + 1
        DataWf.Element(0) = DataWf.Element(0) - TrimOffset
    End If
    Square = (2) ^ CInt(TrimOriginalSize)
    
    
    If InitStateByCapCode = 1 Then
        If b_SetupToBit0 = True Then
            DataWf.Element(0) = DataWf.Element(0) - (Square / ((2) ^ (CInt(TrimOriginalSize) - BitIndex)))
        Else
            DataWf.Element(0) = DataWf.Element(0) + (Square / ((2) ^ (CInt(TrimOriginalSize) - BitIndex)))
        End If
    Else
        If b_SetupToBit0 = True Then
            DataWf.Element(0) = DataWf.Element(0) + (Square / ((2) ^ (CInt(TrimOriginalSize) - BitIndex)))
        Else
            DataWf.Element(0) = DataWf.Element(0) - (Square / ((2) ^ (CInt(TrimOriginalSize) - BitIndex)))
        End If
    End If

    DataWf.Element(0) = DataWf.Element(0) - 1
    DataWf.Element(0) = DataWf.Element(0) + TrimOffset
    CalculateDSP = DataWf.COPY
    CalculateDSP = CalculateDSP.ConvertDataTypeTo(DspLong)
'    DataWf = CalculateDSP.ConvertStreamTo(tldspSerial, TrimOriginalSize + 1, 0, Bit0IsMsb)
    DataWf = CalculateDSP.ConvertStreamTo(tldspSerial, CoverSize, 0, Bit0IsMsb)
    
    
    
''    CalculateSize = DataWf.SampleSize
''    CalculateDSP.CreateConstant 0, 1, DspLong
''    DataWf = DataWf.ConvertDataTypeTo(DspLong)
''    CalculateDSP = DataWf.ConvertStreamTo(tldspParallel, CalculateSize, 0, Bit0IsMsb)
''    CalculateDSP.Element(0) = CalculateDSP.Element(0) - TrimOffset      ' Retrieve original trim code
''    CalculateDSP = CalculateDSP.ConvertDataTypeTo(DspLong)
''    CalculateDSP = CalculateDSP.ConvertStreamTo(tldspSerial, TrimOriginalSize, 0, Bit0IsMsb)

'    If InitStateByCapCode = 0 Then                  '                   |-----------------|-----------------|
'        If b_SetupToBit0 = True Then                ' Code              0                128               256
'            CalculateDSP.Element(BitIndex) = 0      ' Distributed       0000000000000011111111111111111111111
'        Else
'            CalculateDSP.Element(BitIndex) = 1
'        End If
'
'
'    Else
'        If b_SetupToBit0 = True Then                '                   |-----------------|-----------------|
'            CalculateDSP.Element(BitIndex) = 1      ' Code              0                128               256
'        Else                                        ' Distributed       1111111111111000000000000000000000000
'            CalculateDSP.Element(BitIndex) = 0
'        End If
'    End If
'    CalculateDSP = CalculateDSP.ConvertDataTypeTo(DspLong)
'    CalculateDSP = CalculateDSP.ConvertStreamTo(tldspParallel, TrimOriginalSize, 0, Bit0IsMsb)
'    CalculateDSP.Element(0) = CalculateDSP.Element(0) + TrimOffset      ' Addition trimoffset value
'    CalculateDSP = CalculateDSP.ConvertDataTypeTo(DspLong)
'    DataWf = CalculateDSP.ConvertStreamTo(tldspSerial, CalculateSize, 0, Bit0IsMsb)
    
    
End Function

''Public Function DSPWf_Dec2Binary(ByVal InWf As DSPWave, ByVal DataWdth As Long, ByRef OutWf As DSPWave) As Long

    ''OutWf = InWf.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, DataWdth, 0, Bit0IsMsb)

''End Function

Public Function BitWf2Arry_MSB1st(ByVal InWf As DSPWave, ByVal WrdWdth As Integer, _
    ByRef NoOfSamples As Long, ByRef DataWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Convert captured (serial) bit stream to data waveform, Assume MSB->LSB in the bit stream (reversed
    ''''        order may be easily accommodated by adding a switch in the argument list)
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        BitWf2Arry is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        InWf          : DSP Wave (serial) to be converted
    ''''        WrdWdth  : number of bits per word
    ''''        NoOfSamples    : number of samples found in the bit stream
    ''''        DataWf         : converted (parallel) DSP Wave
    ''''
          
    NoOfSamples = InWf.SampleSize
    
    If NoOfSamples Mod WrdWdth <> 0 Then
         Debug.Print vbNewLine & "Bit stream wave size not integer times of the word width." _
            & " Waveform will Be truncated" & vbNewLine
    End If
    
    DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsLsb)
    NoOfSamples = DataWf.SampleSize

End Function

Public Function Split_Dspwave_MSB1st(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
        For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsLsb)
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
End Function
Public Function SetupLinearTrimCodeBit_Linear(ByVal TrimMethod As Boolean, ByRef TrimCode As Double, ByVal b_SetupToBit0 As Boolean, _
ByVal RegSize As Long, ByRef DataWf As DSPWave, ByVal doallFlag As Boolean) As Long
' Dylan Edited 20190529
       Dim i As Integer
    Dim TempDSPWave As New DSPWave
    Dim SizeTemp As Integer
    Dim TrimTemp As Integer
    Dim TotallySize As Integer
    
    SizeTemp = CInt(RegSize) - 1
    TotallySize = CInt(DataWf.SampleSize) - 1
    
    If TrimMethod = True Then                                   ' Increase Linear
        If doallFlag = True Then
            TrimCode = TrimCode + 1
        ElseIf b_SetupToBit0 = False Then
            TrimCode = TrimCode + 1
        End If
    Else                                                        ' Decrease Linear
        If doallFlag = True Then
            TrimCode = TrimCode - 1
        ElseIf b_SetupToBit0 = False Then
            TrimCode = TrimCode - 1
        End If
    End If
    
    
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    TempDSPWave = DataWf.ConvertStreamTo(tldspParallel, TotallySize + 1, 0, Bit0IsMsb)
    TrimTemp = TrimCode + TempDSPWave.Element(0)
    
    If b_SetupToBit0 = False Or doallFlag = True Then
        For i = 0 To SizeTemp
            If TrimTemp <> 0 Then
                If TrimTemp \ (2 ^ (SizeTemp - i)) <> 0 Then
                    DataWf.Element(TotallySize - i) = 1
                Else
                    DataWf.Element(TotallySize - i) = 0
                End If
                If i <> SizeTemp Then
                    If TrimTemp >= (2 ^ (SizeTemp - i)) Then
                        TrimTemp = TrimTemp - (2 ^ (SizeTemp - i))
                    End If
                End If
            Else
                DataWf.Element(TotallySize - i) = 0
            End If
        Next i
    End If
End Function

'TTR,20200423, Oscar
Public Function DSP_CalcMDLLMonotonicityDevideBlock(ByVal InWf As DSPWave, ByRef InWf_DEC_AfterSort As DSPWave, ByRef bDirection As DSPWave, ByRef NumOfDiffVal As DSPWave, ByRef Diff_MaxMin As DSPWave) As Long

    Dim InWf_DEC As New DSPWave
    InWf_DEC = InWf.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 8, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    
    'InWf_DEC = Inwf
    Dim BlockSortArray() As Long
''    BlockSortArray = Array(4, 0, 6, 1, 3, 7, 2, 5)
'    ReDim BlockSortArray(8)
'    BlockSortArray(0) = 4
'    BlockSortArray(1) = 0
'    BlockSortArray(2) = 6
'    BlockSortArray(3) = 1
'    BlockSortArray(4) = 3
'    BlockSortArray(5) = 7
'    BlockSortArray(6) = 2
'    BlockSortArray(7) = 5
'
    Dim i As Long, j As Long, k As Long
'    'Dim InWf_DEC_AfterSort As New DSPWave
'    InWf_DEC_AfterSort.CreateConstant 0, 8, DspLong
'
'    For i = 0 To 7
'        For j = 0 To 7
'            If BlockSortArray(j) = i Then
'                InWf_DEC_AfterSort.Element(i) = InWf_DEC.Element(j)
'            End If
'        Next j
'    Next i

    Dim SortIndex As New DSPWave
    Dim totalsample As Long
    totalsample = InWf_DEC.SampleSize
    SortIndex.CreateConstant 0, totalsample, DspLong
    
    For j = 0 To (totalsample \ 8) - 1
        SortIndex.ElementLite(4 + j * 8) = 0 + j * 8
        SortIndex.ElementLite(0 + j * 8) = 1 + j * 8
        SortIndex.ElementLite(6 + j * 8) = 2 + j * 8
        SortIndex.ElementLite(1 + j * 8) = 3 + j * 8
        SortIndex.ElementLite(3 + j * 8) = 4 + j * 8
        SortIndex.ElementLite(7 + j * 8) = 5 + j * 8
        SortIndex.ElementLite(2 + j * 8) = 6 + j * 8
        SortIndex.ElementLite(5 + j * 8) = 7 + j * 8
    Next j
    
        InWf_DEC_AfterSort = InWf_DEC.Lookup(SortIndex).COPY
        BlockSortArray = InWf_DEC_AfterSort.data
        
    For j = 0 To (totalsample \ 8) - 1
    
        
        For i = 1 To 7
            If BlockSortArray(i + j * 8) - BlockSortArray(i - 1 + j * 8) > 0 Then
                bDirection.Element(j) = 0
            End If
        Next i
        
        Dim TempVal As Long
        
        If bDirection.Element(j) = 0 Then
            For i = 0 To 7
                For k = i + 1 To 6
                    If BlockSortArray(i + j * 8) < BlockSortArray(k + j * 8) Then
                        TempVal = BlockSortArray(i + j * 8)
                        BlockSortArray(i + j * 8) = BlockSortArray(k + j * 8)
                        BlockSortArray(k + j * 8) = TempVal
                    End If
                Next k
            Next i
        End If
        
        For i = 0 To (7 - 1)
        
            If BlockSortArray(i + 1 + j * 8) < BlockSortArray(i + j * 8) Then NumOfDiffVal.Element(j) = NumOfDiffVal.Element(j) + 1
        
        Next i
           
        Diff_MaxMin.Element(j) = InWf_DEC.Select(0 + j * 8, 1, 8).CalcMaximumValue - InWf_DEC.Select(0 + j * 8, 1, 8).CalcMinimumValue
    Next j

End Function
'TTR,20200423, Oscar
Public Function Split_Tolerance_DSPWave(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf_Final As DSPWave) As Long
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    Dim OutWf As New DSPWave: OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    Dim width_Wf_Concatenate As New DSPWave: width_Wf_Concatenate.CreateConstant 0, width_Wf.SampleSize / 2, DspDouble
    Dim j As Long: j = 0
    Dim OutWf_Concatenate As New DSPWave: OutWf_Concatenate.CreateConstant 0, width_Wf.SampleSize / 2, DspDouble
    Dim Threshold_P_Flag As Boolean: Threshold_P_Flag = False
    Dim Threshold_N_Flag As Boolean: Threshold_N_Flag = False
    Dim Threshold_P As New DSPWave: Threshold_P.CreateConstant 999, 2, DspDouble
    Dim Threshold_N As New DSPWave: Threshold_N.CreateConstant 999, 2, DspDouble
    Dim Threshold_P_Found As New DSPWave: Threshold_P_Found.CreateConstant 0, 1, DspDouble
    Dim Threshold_N_Found As New DSPWave: Threshold_N_Found.CreateConstant 0, 1, DspDouble
    Dim Threshold_P_Found_Count As Long: Threshold_P_Found_Count = 0
    Dim Threshold_N_Found_Count As Long: Threshold_N_Found_Count = 0
    Dim Threshold_Mapping As New DSPWave: Threshold_Mapping.CreateConstant 0, 16, DspDouble
    Threshold_Mapping.Element(0) = -0.2
    Threshold_Mapping.Element(1) = -0.15
    Threshold_Mapping.Element(2) = -0.1
    Threshold_Mapping.Element(3) = -0.05
    Threshold_Mapping.Element(4) = 0.05
    Threshold_Mapping.Element(5) = 0.1
    Threshold_Mapping.Element(6) = 0.15
    Threshold_Mapping.Element(7) = 0.2
    current_loc = 0
    For i = 0 To width_Wf.SampleSize - 1 Step 2
        width_Wf_Concatenate.Element(j) = width_Wf.Element(i) + width_Wf.Element(i + 1)
        j = j + 1
    Next i
    For i = 0 To width_Wf_Concatenate.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf_Concatenate.Element(i)).ConvertStreamTo(tldspParallel, width_Wf_Concatenate.Element(i), 0, Bit0IsMsb)
        OutWf_Concatenate.Element(i) = dec_wave.Element(0)
        If i < width_Wf_Concatenate.SampleSize / 2 And OutWf_Concatenate.Element(i) = 0 And Threshold_P_Flag = False Then
            Threshold_P.Element(0) = Threshold_Mapping.Element(i)
            Threshold_P_Flag = True
        ElseIf i < width_Wf_Concatenate.SampleSize / 2 And OutWf_Concatenate.Element(i) = 0 And Threshold_P_Flag = True Then
            Threshold_P.Element(1) = Threshold_Mapping.Element(i)
        ElseIf OutWf_Concatenate.Element(i) = 0 And Threshold_N_Flag = False Then
            Threshold_N.Element(0) = Threshold_Mapping.Element(i - width_Wf_Concatenate.SampleSize / 2)
            Threshold_N_Flag = True
        ElseIf OutWf_Concatenate.Element(i) = 0 And Threshold_N_Flag = True Then
            Threshold_N.Element(1) = Threshold_Mapping.Element(i - width_Wf_Concatenate.SampleSize / 2)
        Else 'Do nothing '20230601
        End If
        If i < width_Wf_Concatenate.SampleSize / 2 And OutWf_Concatenate.Element(i) > 0 Then
            Threshold_P_Found_Count = Threshold_P_Found_Count + 1
        ElseIf OutWf_Concatenate.Element(i) > 0 Then
            Threshold_N_Found_Count = Threshold_N_Found_Count + 1
        Else 'Do nothing '20230601
        End If
        current_loc = current_loc + width_Wf_Concatenate.Element(i)
    Next i
    If Threshold_P_Found_Count > 0 And Threshold_P_Flag Then Threshold_P_Found.Element(0) = 1
    If Threshold_N_Found_Count > 0 And Threshold_N_Flag Then Threshold_N_Found.Element(0) = 1
    OutWf_Final = OutWf.Concatenate(Threshold_P).Concatenate(Threshold_P_Found).Concatenate(Threshold_N).Concatenate(Threshold_N_Found)
End Function
'TTR,20200423, Oscar
Public Function Split_Dspwave_ADC_Convert_Avg(ByVal SourceBitStrmWf As DSPWave, ByRef width_Wf As DSPWave, ByRef OutWf As DSPWave, ByVal AvgCount As Long, ByVal ref As Long, ByRef AvgDataPos As DSPWave, ByVal Gap As Long, ByVal blockNum As Long, ByVal Addvalue As Double, ByVal MinusValue As Double, _
                                                ByRef OutWf_RefAverage As DSPWave, ByRef OutWf_CalcAverage As DSPWave) As Long

Dim temp_RefferanceCode_DEC As Double
Dim ADC_code_average_DEC As Double
Dim ADC_code_DEC As Double
Dim BlockIndex As Long
Dim AvgPos_Arr() As Long
Dim AvgCalcNumber As Long
Dim i As Long: Dim j As Long

AvgCalcNumber = AvgDataPos.SampleSize

Call Split_Dspwave(SourceBitStrmWf, width_Wf, OutWf)

AvgPos_Arr = AvgDataPos.data

For BlockIndex = 0 To blockNum - 1
    temp_RefferanceCode_DEC = 0
    For i = 0 To AvgCount - 1
            temp_RefferanceCode_DEC = temp_RefferanceCode_DEC + OutWf.Element(ref + i + BlockIndex * Gap)
    Next i
    temp_RefferanceCode_DEC = temp_RefferanceCode_DEC / AvgCount
    
    For j = 0 To AvgCalcNumber - 1 Step AvgCount
            ADC_code_average_DEC = 0
            For i = 0 To AvgCount - 1
                ADC_code_DEC = (((OutWf.Element(AvgPos_Arr(i + j) + BlockIndex * Gap) + Addvalue - temp_RefferanceCode_DEC) / 256) * 0.5 * MinusValue) + (0.25 * MinusValue)
                ADC_code_average_DEC = ADC_code_average_DEC + ADC_code_DEC
            Next i
            ADC_code_average_DEC = ADC_code_average_DEC / AvgCount
            OutWf_CalcAverage.Element((BlockIndex) * AvgCalcNumber + j) = ADC_code_average_DEC
    Next j
    
    OutWf_RefAverage.Element(BlockIndex) = temp_RefferanceCode_DEC
Next BlockIndex

End Function

Public Function Split_Dspwave_AUS_TTR(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, Outwf_T1 As DSPWave, storeDSP As DSPWave) As Long
                                                                                                                                                                                                                                                               
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim temp_dsp As New DSPWave
    Dim Count As Long
    current_loc = 0
    Count = 0
                                                                                                                                                                                                                                                                '' 20170705 - Solve for over 32 bits to decimal
    temp_dsp.CreateConstant 0, width_Wf.SampleSize, DspDouble
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    Outwf_T1.CreateConstant 0, width_Wf.SampleSize + 1, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
                                                                                                                                                                                                                                                              
    Outwf_T1.Select(0, 1, OutWf.SampleSize).Replace OutWf.Select(0, 1, -1)
    
    ' 1.2 Calculate desired index
    Dim target_index As Long: target_index = 999
    Dim target_gap As Double: target_gap = 2013144
    For i = 0 To OutWf.SampleSize - 2
        If (OutWf.Element(i) - 3333) >= 0 Then
            target_gap = Abs(OutWf.Element(i) - 3333)
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next
    
    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_gap = 2013144
        For i = 0 To OutWf.SampleSize - 2
            If Abs(OutWf.Element(i) - 3333) < target_gap Then
                target_gap = Abs(OutWf.Element(i) - 3333)
                target_index = i
            End If
        Next
    End If
'
    'OutWf.Element(OutWf.SampleSize - 1) = target_index
    
    Dim TempVal As Long
    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T1.Element(OutWf.SampleSize) = TempVal
    
    For i = 0 To CLng((storeDSP.SampleSize)) - 1
        storeDSP.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            storeDSP.Element(i) = 0
        End If
    Next i
                                                                                                                                                                                                                                                               
End Function

Public Function Split_Dspwave_CIO3PLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, _
Delta_Value As DSPWave, ByVal calibration_target_value_low As Long, ByVal calibration_target_value_high As Long, _
ByVal start_search As Long, Bin_Target_low As DSPWave, Bin_Target_high As DSPWave, ByVal Interspersed_boolen As Boolean, _
ByVal z As Long, ByVal Lane_Sum As Long, ByVal Each_lane_size As Long, ByVal offset As Long) As Long


    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim Interspersed_DSP As New DSPWave
    Dim target_var_low As Long
    Dim target_var_high As Long
    
    current_loc = 0
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
            dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
            OutWf.Element(i) = dec_wave.Element(0)
            current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    
     If Interspersed_boolen = True Then
        Interspersed_DSP = OutWf.Select(z, Lane_Sum, Each_lane_size).COPY
       Else
        Interspersed_DSP = OutWf.Select(z * Each_lane_size, 1, Each_lane_size).COPY
     End If
    
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    

    For i = start_search To CLng(Interspersed_DSP.SampleSize) - 1
        calc_data.Element(i - start_search) = Interspersed_DSP.Element(i)
    Next i
    ''' compare the target


    temp_delta_value = 9999
                 
    For k = start_search To CLng((Interspersed_DSP.SampleSize)) - 1
        Delta_Value.Element(k) = Abs(calibration_target_value_low - calc_data.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var_low = k
        End If
    Next k
    
    temp_delta_value = 9999
    
    For k = start_search To CLng((Interspersed_DSP.SampleSize)) - 1
        Delta_Value.Element(k) = Abs(calibration_target_value_high - calc_data.Element(k - start_search))
        'search min delta
        If Delta_Value.Element(k) < temp_delta_value Then
            temp_delta_value = Delta_Value.Element(k)
            target_var_high = k
        End If
    Next k

    target_var_low = target_var_low + offset
    target_var_high = target_var_high + offset
    ''''''''''''' Decimal to Binary dspwave for dictionary'''''''''''''''
    
    Dim TempVal As Long
    
    TempVal = target_var_low
    For i = 0 To CLng((Bin_Target_low.SampleSize)) - 1
        Bin_Target_low.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_low.Element(i) = 0
        End If
    Next i
    
    TempVal = target_var_high
    For i = 0 To CLng((Bin_Target_high.SampleSize)) - 1
        Bin_Target_high.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_high.Element(i) = 0
        End If
    Next i
    
End Function


Public Function Split_Dspwave_ASGMTR_Freq_Comp_Decomp(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, ByVal MTRS_BYP_Comp_Decomp As Boolean, _
                                    ByVal MTRSNS_Matrix_ROT_Row As Long, ByVal MTRSNS_Matrix_ROV_Row As Long, ByVal MTRSNS_Matrix_ROT_Column As Long, ByVal MTRSNS_Matrix_ROV_Column As Long, _
                                    ByVal DSP_ROT_a_max_min As DSPWave, ByVal DSP_ROV_a_max_min As DSPWave, ByVal SensorNum As Long, _
                                    ByVal DSP_MetrologySense_ROT_Matrix As DSPWave, ByVal DSP_MetrologySense_ROV_Matrix As DSPWave, _
                                    ByRef MetrologySense_ROT_Frequency As DSPWave, ByRef MetrologySense_ROV_Frequency As DSPWave, _
                                    ByRef a1 As DSPWave, ByRef a2 As DSPWave, ByRef a1_Compression As DSPWave, ByRef a2_Compression As DSPWave, _
                                    ByRef a1_Compression_eFuse As DSPWave, ByRef a2_Compression_eFuse As DSPWave, _
                                    ByRef MetrologySense_ROT_Frequency_DeCompression As DSPWave, ByRef MetrologySense_ROV_Frequency_DeCompression As DSPWave, _
                                    ByRef MetrologySense_ROT_Frequency_Error As DSPWave, ByRef MetrologySense_ROV_Frequency_Error As DSPWave, _
                                    ByRef MetrologySense_ROT_Frequency_Monotonicity As DSPWave, ByRef MetrologySense_ROV_Frequency_Monotonicity As DSPWave) As Long

    Dim MetrologySense_Frequency As New DSPWave
    Dim SensorIndex As Long
    Dim i As Long
    
    Call Split_Dspwave(InWf, width_Wf, OutWf)
    
    MetrologySense_Frequency = OutWf.Multiply(50000)
    
    MetrologySense_ROT_Frequency = MetrologySense_Frequency.Select(0, 2).COPY
    MetrologySense_ROV_Frequency = MetrologySense_Frequency.Select(1, 2).COPY
    
    If MTRS_BYP_Comp_Decomp = True Then Exit Function 'Tonga
    
    Dim a1_max As New DSPWave
    Dim a1_min As New DSPWave
    Dim a2_max As New DSPWave
    Dim a2_min As New DSPWave
    
    
    Dim L_ROT_Frequency As New DSPWave
    Dim L_ROV_Frequency As New DSPWave
    Dim L_a1 As New DSPWave
    Dim L_a2 As New DSPWave
    Dim L_a1_Compression As New DSPWave
    Dim L_a2_Compression As New DSPWave
    Dim L_a1_Compression_eFuse As New DSPWave
    Dim L_a2_Compression_eFuse As New DSPWave
    Dim L_Rtn_a1_Compression As New DSPWave
    Dim L_Rtn_a2_Compression As New DSPWave
    Dim L_a1_DeCompression As New DSPWave
    Dim L_a2_DeCompression As New DSPWave
    Dim L_ROT_Frequency_DeCompression As New DSPWave
    Dim L_ROV_Frequency_DeCompression As New DSPWave
    Dim L_ROT_Frequency_Error As New DSPWave
    Dim L_ROV_Frequency_Error As New DSPWave
    Dim L_ROT_Frequency_Monotonicity As New DSPWave
    Dim L_ROV_Frequency_Monotonicity As New DSPWave
    
    
    a1_max = DSP_ROT_a_max_min.Select(0, 2).COPY
    a1_min = DSP_ROT_a_max_min.Select(1, 2).COPY
    a2_max = DSP_ROV_a_max_min.Select(0, 2).COPY
    a2_min = DSP_ROV_a_max_min.Select(1, 2).COPY
    
    For SensorIndex = 0 To SensorNum - 1
    
        L_ROT_Frequency = MetrologySense_ROT_Frequency.divide(10 ^ 9).COPY.Select(SensorIndex, SensorNum, MTRSNS_Matrix_ROT_Column).COPY
        L_ROV_Frequency = MetrologySense_ROV_Frequency.divide(10 ^ 9).COPY.Select(SensorIndex, SensorNum, MTRSNS_Matrix_ROV_Column).COPY
        
        L_ROT_Frequency_Monotonicity = L_ROT_Frequency.Select(1).Subtract(L_ROT_Frequency.Select(0))
        L_ROV_Frequency_Monotonicity = L_ROV_Frequency.Select(1).Subtract(L_ROV_Frequency.Select(0))
        
        L_a1 = DSP_MetrologySense_ROT_Matrix.MatrixMultiply(MTRSNS_Matrix_ROT_Row, MTRSNS_Matrix_ROT_Column, L_ROT_Frequency).COPY
        L_a2 = DSP_MetrologySense_ROV_Matrix.MatrixMultiply(MTRSNS_Matrix_ROV_Row, MTRSNS_Matrix_ROV_Column, L_ROV_Frequency).COPY
    
        L_a1_Compression = L_a1.Subtract(a1_min).divide(a1_max.Subtract(a1_min)).COPY
        L_a2_Compression = L_a2.Subtract(a2_min).divide(a2_max.Subtract(a2_min)).COPY
        
        L_a1_Compression_eFuse.CreateConstant -1, MTRSNS_Matrix_ROT_Row, DspLong 'Format(,0) Be aware of (2 ^ 31 -1) is the maximum value, added by Sean, +886 965507602
        L_a2_Compression_eFuse.CreateConstant -1, MTRSNS_Matrix_ROV_Row, DspLong
        
        For i = 0 To MTRSNS_Matrix_ROT_Row - 1
            If i = 0 Then
                If L_a1_Compression.Element(i) >= 1 Then
                    L_a1_Compression_eFuse.Element(i) = 2 ^ 15 - 1
                Else
                    L_a1_Compression_eFuse.Element(i) = L_a1_Compression.Element(i) * 2 ^ 15
                End If
            Else
                If L_a1_Compression.Element(i) >= 1 Then
                    L_a1_Compression_eFuse.Element(i) = 2 ^ 14 - 1
                Else
                    L_a1_Compression_eFuse.Element(i) = L_a1_Compression.Element(i) * 2 ^ 14
                End If
            End If
        Next i
        For i = 0 To MTRSNS_Matrix_ROV_Row - 1
            If i = 0 Then
                If L_a2_Compression.Element(i) >= 1 Then
                    L_a2_Compression_eFuse.Element(i) = 2 ^ 15 - 1
                Else
                    L_a2_Compression_eFuse.Element(i) = L_a2_Compression.Element(i) * 2 ^ 15
                End If
            Else
                If L_a2_Compression.Element(i) >= 1 Then
                    L_a2_Compression_eFuse.Element(i) = 2 ^ 14 - 1
                Else
                    L_a2_Compression_eFuse.Element(i) = L_a2_Compression.Element(i) * 2 ^ 14
                End If
            End If
        Next i
        
        L_Rtn_a1_Compression.CreateConstant -1, MTRSNS_Matrix_ROT_Row, DspDouble
        L_Rtn_a2_Compression.CreateConstant -1, MTRSNS_Matrix_ROV_Row, DspDouble
        
        For i = 0 To MTRSNS_Matrix_ROT_Row - 1
            If i = 0 Then
                L_Rtn_a1_Compression.Element(i) = L_a1_Compression_eFuse.Element(i) / (2 ^ 15)
            Else
                L_Rtn_a1_Compression.Element(i) = L_a1_Compression_eFuse.Element(i) / (2 ^ 14)
            End If
        Next i
        For i = 0 To MTRSNS_Matrix_ROV_Row - 1
            If i = 0 Then
                L_Rtn_a2_Compression.Element(i) = L_a2_Compression_eFuse.Element(i) / (2 ^ 15)
            Else
                L_Rtn_a2_Compression.Element(i) = L_a2_Compression_eFuse.Element(i) / (2 ^ 14)
            End If
        Next i
        
        L_a1_DeCompression = L_Rtn_a1_Compression.Multiply(a1_max.Subtract(a1_min)).Add(a1_min).COPY
        L_a2_DeCompression = L_Rtn_a2_Compression.Multiply(a2_max.Subtract(a2_min)).Add(a2_min).COPY
        L_ROT_Frequency_DeCompression = DSP_MetrologySense_ROT_Matrix.MatrixTranspose(MTRSNS_Matrix_ROT_Row).MatrixMultiply(MTRSNS_Matrix_ROT_Column, MTRSNS_Matrix_ROT_Row, L_a1_DeCompression).COPY
        L_ROV_Frequency_DeCompression = DSP_MetrologySense_ROV_Matrix.MatrixTranspose(MTRSNS_Matrix_ROV_Row).MatrixMultiply(MTRSNS_Matrix_ROV_Column, MTRSNS_Matrix_ROV_Row, L_a2_DeCompression).COPY
        L_ROT_Frequency_DeCompression = L_ROT_Frequency_DeCompression.Multiply(10 ^ 9).COPY
        L_ROV_Frequency_DeCompression = L_ROV_Frequency_DeCompression.Multiply(10 ^ 9).COPY

        

        'Return Value
        If SensorIndex = 0 Then
            a1 = L_a1.COPY
            a2 = L_a2.COPY
            a1_Compression = L_a1_Compression.COPY
            a2_Compression = L_a2_Compression.COPY
            a1_Compression_eFuse = L_a1_Compression_eFuse.COPY
            a2_Compression_eFuse = L_a2_Compression_eFuse.COPY
            MetrologySense_ROT_Frequency_DeCompression = L_ROT_Frequency_DeCompression.COPY
            MetrologySense_ROV_Frequency_DeCompression = L_ROV_Frequency_DeCompression.COPY
            MetrologySense_ROT_Frequency_Error = L_ROT_Frequency_DeCompression.Subtract(L_ROT_Frequency.Multiply(10 ^ 9)).divide(L_ROT_Frequency_DeCompression).COPY
            MetrologySense_ROV_Frequency_Error = L_ROV_Frequency_DeCompression.Subtract(L_ROV_Frequency.Multiply(10 ^ 9)).divide(L_ROV_Frequency_DeCompression).COPY
            
            MetrologySense_ROT_Frequency_Monotonicity = L_ROT_Frequency_Monotonicity.Multiply(10 ^ 9).COPY
            MetrologySense_ROV_Frequency_Monotonicity = L_ROV_Frequency_Monotonicity.Multiply(10 ^ 9).COPY
            
        Else
            a1 = a1.Concatenate(L_a1).COPY
            a2 = a2.Concatenate(L_a2).COPY
            a1_Compression = a1_Compression.Concatenate(L_a1_Compression).COPY
            a2_Compression = a2_Compression.Concatenate(L_a2_Compression).COPY
            a1_Compression_eFuse = a1_Compression_eFuse.Concatenate(L_a1_Compression_eFuse).COPY
            a2_Compression_eFuse = a2_Compression_eFuse.Concatenate(L_a2_Compression_eFuse).COPY
            MetrologySense_ROT_Frequency_DeCompression = MetrologySense_ROT_Frequency_DeCompression.Concatenate(L_ROT_Frequency_DeCompression).COPY
            MetrologySense_ROV_Frequency_DeCompression = MetrologySense_ROV_Frequency_DeCompression.Concatenate(L_ROV_Frequency_DeCompression).COPY
            MetrologySense_ROT_Frequency_Error = MetrologySense_ROT_Frequency_Error.Concatenate(L_ROT_Frequency_DeCompression.Subtract(L_ROT_Frequency.Multiply(10 ^ 9)).divide(L_ROT_Frequency_DeCompression)).COPY
            MetrologySense_ROV_Frequency_Error = MetrologySense_ROV_Frequency_Error.Concatenate(L_ROV_Frequency_DeCompression.Subtract(L_ROV_Frequency.Multiply(10 ^ 9)).divide(L_ROV_Frequency_DeCompression)).COPY
            
            MetrologySense_ROT_Frequency_Monotonicity = MetrologySense_ROT_Frequency_Monotonicity.Concatenate(L_ROT_Frequency_Monotonicity.Multiply(10 ^ 9)).COPY
            MetrologySense_ROV_Frequency_Monotonicity = MetrologySense_ROV_Frequency_Monotonicity.Concatenate(L_ROV_Frequency_Monotonicity.Multiply(10 ^ 9)).COPY
            
        End If
                
    Next SensorIndex
    
    
End Function

Public Function Split_Dspwave_AUS_TTR_Charz(ByRef InWf As DSPWave, ByRef width_Wf As DSPWave, ByRef OutWf As DSPWave, ByRef Outwf_T1 As DSPWave, _
ByRef sda_measuredata As DSPWave, ByRef storeDSP As DSPWave, ByVal TargetFreq As Double, ByVal start_search As Long, ByVal search_width As Long, _
ByVal Interspersed_boolen As Boolean, ByVal z As Long, ByVal Lane_Sum As Long, ByRef Trim_target_final As Long) As Long
                                                                                                                                                                                                                                                               
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim i As Long
    Dim j As Long
    Dim Count As Long
    Dim current_loc As Long
    Dim dec_wave As New DSPWave
    Dim DSPDiscard As New DSPWave
    Dim temp_dsp() As New DSPWave
    Dim lane_decide As New DSPWave
    Dim Outwf_T1_Temp As New DSPWave
    ReDim temp_dsp(2)
    
    Count = 0
    current_loc = 0
        
    DSPDiscard.CreateConstant 9999, 3, DspLong
    lane_decide.CreateConstant 9999, 3, DspDouble
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    Outwf_T1_Temp.CreateConstant 0, sda_measuredata.SampleSize * Lane_Sum, DspLong
    
    For i = 0 To 2
        temp_dsp(i).CreateConstant 9999, (sda_measuredata.SampleSize / 3), DspDouble
    Next i
    
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        If width_Wf.Element(i) = search_width Then
            Outwf_T1_Temp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
        End If
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    If Interspersed_boolen = True Then ' CAPTURE count is not inculde buf and unbuf
        Outwf_T1 = Outwf_T1_Temp.Select(z, Lane_Sum, sda_measuredata.SampleSize).COPY   ' (number of lane + VRO offset, jump how may capture, capture size)
        '---------------------------------------------------------------------------------------------
        ' For Example. z = lane sweep number , Lane_Sum = 3
        ' FCNT_GP_F0
        ' FCNT_ST_F0
        ' FCNT_EX_F0
        ' FCNT_GP_F1
        ' FCNT_ST_F1
        ' FCNT_EX_F1
        '---------------------------------------------------------------------------------------------
    Else
        Outwf_T1 = Outwf_T1_Temp.Select(z * sda_measuredata.SampleSize, 1, sda_measuredata.SampleSize).COPY  ' CAPTURE is not inculde buf and unbuf
        '---------------------------------------------------------------------------------------------
        ' For Example. z = lane sweep number
        ' FCNT_GP_F0
        ' FCNT_GP_F1
        ' FCNT_ST_F0
        ' FCNT_ST_F1
        ' FCNT_EX_F0
        ' FCNT_EX_F1
        '---------------------------------------------------------------------------------------------
    End If
    
    Dim target_index As Long: target_index = 999
    Dim target_gap As Double: target_gap = 2013144

    For j = 0 To 2
        For i = 0 To CLng(sda_measuredata.SampleSize / 3) - 1
            If sda_measuredata.Element(i + CLng(sda_measuredata.SampleSize / 3) * j) >= 0.3 Then
                If Outwf_T1.Element(i + CLng(sda_measuredata.SampleSize / 3) * j) - TargetFreq >= 0 Then
                    temp_dsp(j).Element(i) = Abs(Outwf_T1.Element(i + CLng(sda_measuredata.SampleSize / 3) * j) - TargetFreq)
                    DSPDiscard.Element(j) = 1
                End If
            End If
        Next i
    Next j
    
    If DSPDiscard.CalcMinimumValue = 9999 Then
        For j = 0 To 2
            For i = 0 To CLng(sda_measuredata.SampleSize / 3) - 1
                If sda_measuredata.Element(i + CLng(sda_measuredata.SampleSize / 3) * j) >= 0.3 Then
                    If Abs(Outwf_T1.Element(i + CLng(sda_measuredata.SampleSize / 3) * j) - TargetFreq) < target_gap Then
                        temp_dsp(j).Element(i) = Abs(Outwf_T1.Element(i + CLng(sda_measuredata.SampleSize / 3) * j) - TargetFreq)
                        DSPDiscard.Element(j) = 2
                    End If
                End If
            Next i
        Next j
    End If
    
    Dim lane_min As Double
    Dim Disable_records As Boolean: Disable_records = True
    Dim en_cap As Long
    Dim Record_targetVro As Double
    
    For j = 0 To 2
        If temp_dsp(j).CalcMinimumValue <> 9999 Then
            lane_decide.Element(j) = temp_dsp(j).CalcMinimumValue
            Disable_records = False
        End If
    Next
   
    If Disable_records = False Then
'''''        lane_min = lane_decide.CalcMinimumValue
        For j = 0 To 2
'''''            If lane_decide.Element(j) = lane_min And DSPDiscard.Element(j) <> 9999 Then
            If DSPDiscard.Element(j) <> 9999 Then
               en_cap = j
               Exit For
            End If
        Next j
        
        Record_targetVro = temp_dsp(en_cap).CalcMinimumValue
        
        For i = 0 To CLng(sda_measuredata.SampleSize / 3) - 1
            If temp_dsp(en_cap).Element(i) = Record_targetVro Then
               target_index = i
               Exit For
            End If
        Next i
    End If
    
    dec_wave.CreateConstant 0, 1, DspLong
    If en_cap = 2 Then en_cap = en_cap + 1
    dec_wave.Element(0) = target_index + en_cap * 8
    storeDSP = dec_wave.ConvertStreamTo(tldspSerial, 5, 0, Bit0IsMsb)
    Trim_target_final = dec_wave.Element(0)
                                                                                                                                                                                    
End Function

Public Function Calc_EquationTrim_DSP(ByRef Result_Code As Long, ByRef Result_DiffVal As Double, ByRef Result_CodeBinDsp As DSPWave, ByRef Result_CodeDecDsp As DSPWave, ByRef CalculatedDsp As DSPWave, _
                                      ByVal TrimTarget As Double, ByVal InputDicVal As Double, ByVal FuseBitNum As Long, ByVal TrimCodeSize As Long, ByVal TrimCodeEqDsp_Coef As DSPWave) As Long

    Dim TrimCodeEqDsp As New DSPWave
    Dim tempLong As Long
    
    TrimCodeEqDsp.CreatePolynomial TrimCodeEqDsp_Coef, TrimCodeSize + 1
    CalculatedDsp = TrimCodeEqDsp.Multiply(InputDicVal).Subtract(TrimTarget).Abs.COPY
    Result_DiffVal = CalculatedDsp.CalcMinimumValue(tempLong)
    Result_Code = tempLong
    Result_CodeDecDsp.Element(0) = Result_Code
    Result_CodeBinDsp = Result_CodeDecDsp.ConvertStreamTo(tldspSerial, FuseBitNum, 0, Bit0IsMsb)
        
End Function

Public Function Split_Dspwave_MSB(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, MSB_Wf As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
        For i = 0 To width_Wf.SampleSize - 1
                If MSB_Wf.Element(i) = 0 Then
                    dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
                Else
                    dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsLsb)
                End If
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
End Function


Public Function Split_Dspwave_Reverse(ByVal InWf As DSPWave, OutWf As DSPWave) As Long
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
'    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    dec_wave.CreateConstant 0, 1, DspDouble
    OutWf.CreateConstant 0, InWf.SampleSize, DspDouble
    For i = 0 To InWf.SampleSize - 1
        dec_wave.Element(0) = InWf.Element(InWf.SampleSize - 1 - i)
        OutWf.Element(i) = dec_wave.Element(0)
    Next i
End Function


Public Function DSP_CalcMDLLMonotonicityDevideBlockSEG(ByVal InWf As DSPWave, ByRef InWf_Bin2DEC_Sort As DSPWave, ByRef Diff_MaxMin As DSPWave, ByRef bDirection As DSPWave, ByRef Mean_Val As DSPWave, ByRef bUniqueDirection As DSPWave, ByRef SumVal As DSPWave) As Long
    'Update from Staten @William 211101
    Dim i As Long
    Dim j As Long
    Dim totalsample As Long
    Dim totalGrpsample As Long
    Dim InWf_DecGrp_SortAry() As Long
    
    Dim InWf_Dec2Bin As New DSPWave
    
    Dim InWf_Bin_SortAry() As New DSPWave
    
    InWf_Dec2Bin = InWf.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    
    totalsample = 96 ''InWf_Dec2Bin.SampleSize \ 2 ''96 (32 * 3)
    totalGrpsample = InWf_Dec2Bin.SampleSize \ 12 ''8
    
    ReDim InWf_Bin_SortAry(totalGrpsample - 1)
    ReDim InWf_DecGrp_SortAry(totalGrpsample - 1)
    
    For j = 0 To (totalGrpsample \ 8) - 1
        InWf_Bin_SortAry(4 + j * 8) = InWf_Dec2Bin.Select(0 + j * totalsample, 1, 9).COPY
        InWf_Bin_SortAry(0 + j * 8) = InWf_Dec2Bin.Select(10 + j * totalsample, 1, 9).COPY
        InWf_Bin_SortAry(6 + j * 8) = InWf_Dec2Bin.Select(20 + j * totalsample, 1, 9).COPY
        InWf_Bin_SortAry(1 + j * 8) = InWf_Dec2Bin.Select(32 + j * totalsample, 1, 9).COPY
        InWf_Bin_SortAry(3 + j * 8) = InWf_Dec2Bin.Select(42 + j * totalsample, 1, 9).COPY
        InWf_Bin_SortAry(7 + j * 8) = InWf_Dec2Bin.Select(52 + j * totalsample, 1, 9).COPY
        InWf_Bin_SortAry(2 + j * 8) = InWf_Dec2Bin.Select(64 + j * totalsample, 1, 9).COPY
        InWf_Bin_SortAry(5 + j * 8) = InWf_Dec2Bin.Select(74 + j * totalsample, 1, 9).COPY
    Next j
    
    For j = 0 To totalGrpsample - 1
        InWf_Bin_SortAry(j) = InWf_Bin_SortAry(j).ConvertStreamTo(tldspParallel, 9, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
        InWf_DecGrp_SortAry(j) = InWf_Bin_SortAry(j).Element(0)
    Next

    InWf_Bin2DEC_Sort.data = InWf_DecGrp_SortAry
    
    For j = 0 To (totalGrpsample \ 8) - 1
        bDirection.Element(j) = 1
        For i = 0 To 6
'            If (InWf_DecGrp_SortAry(i + j * 8) + 1 = InWf_DecGrp_SortAry(i + 1 + j * 8)) Or (InWf_DecGrp_SortAry(i + j * 8) - 1 = InWf_DecGrp_SortAry(i + 1 + j * 8)) Then
''                bDirection.Element(j) = 2
'            ElseIf (InWf_DecGrp_SortAry(i + j * 8) = InWf_DecGrp_SortAry(i + 1 + j * 8)) Then
            
''                bDirection.Element(j) = -2
            If (InWf_DecGrp_SortAry(i + j * 8) < InWf_DecGrp_SortAry(i + 1 + j * 8)) Then
                bDirection.Element(j) = 0
                Exit For
                'Debug.Print "set decrease bit to 0 "
            Else
            End If
        Next i
        
        bUniqueDirection.Element(j) = 1
        For i = 0 To 6
            'If bUniqueDirection.Element(j) = 1 Then
'                If InWf_DecGrp_SortAry(i + j * 8) = InWf_DecGrp_SortAry(i + 1 + j * 8) Then
'                    bUniqueDirection.Element(j) = 1
'                ElseIf InWf_DecGrp_SortAry(i + j * 8) = InWf_DecGrp_SortAry(i + 1 + j * 8) + 1 Then
'                    bUniqueDirection.Element(j) = 2
'                Else
'                    bUniqueDirection.Element(j) = -2
'                End If
'            ElseIf bUniqueDirection.Element(j) = 2 Then
'                If InWf_DecGrp_SortAry(i + j * 8) = InWf_DecGrp_SortAry(i + 1 + j * 8) Then
'                    bUniqueDirection.Element(j) = 2
'                Else
'                    bUniqueDirection.Element(j) = -2
'            End If
            If InWf_DecGrp_SortAry(i + j * 8) <> InWf_DecGrp_SortAry(i + 1 + j * 8) Then
                bUniqueDirection.Element(j) = bUniqueDirection.Element(j) + 1
            Else
            End If
        Next i
        
        
        If Abs(InWf_Bin2DEC_Sort.Select(0 + j * 8, 1, 8).CalcMaximumValue - InWf_Bin2DEC_Sort.Select(0 + j * 8, 1, 8).CalcMinimumValue) > 15 Then
            bUniqueDirection.Element(j) = -1
        End If
        Diff_MaxMin.Element(j) = InWf_Bin2DEC_Sort.Select(0 + j * 8, 1, 8).CalcMaximumValue - InWf_Bin2DEC_Sort.Select(0 + j * 8, 1, 8).CalcMinimumValue
        SumVal.Element(j) = InWf_Bin2DEC_Sort.Select(0 + j * 8, 1, 8).CalcSum
        Mean_Val.Element(j) = InWf_Bin2DEC_Sort.Select(0 + j * 8, 1, 8).CalcMean
    Next j

End Function


Public Function Split_Dspwave_LPDDR5PLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, Outwf_T1 As DSPWave, Outwf_T2 As DSPWave, Outwf_T3 As DSPWave, Outwf_T4 As DSPWave, storeDSP1 As DSPWave, StoreDSP2 As DSPWave, storeDSP3 As DSPWave, storeDSP4 As DSPWave) As Long
    
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim temp_dsp As New DSPWave
    Dim Count As Long
    Dim BiasSize As Long: BiasSize = width_Wf.SampleSize / 4
    Dim STPos As Long
    
    current_loc = 0
    Count = 0
    
    temp_dsp.CreateConstant 0, width_Wf.SampleSize, DspDouble
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    
    ' Fetch capture to array for later trimming
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        If width_Wf.Element(i) = 16 Then
            temp_dsp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
        End If
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    'STPos = Count
    STPos = Count / 4
    
    Outwf_T1.CreateConstant 0, BiasSize, DspDouble
    Outwf_T2.CreateConstant 0, BiasSize, DspDouble
    Outwf_T3.CreateConstant 0, BiasSize, DspDouble
    Outwf_T4.CreateConstant 0, BiasSize, DspDouble
    
    Outwf_T1.Select(0, 1, BiasSize).Replace temp_dsp.Select(0, 1, BiasSize)
    Outwf_T2.Select(0, 1, BiasSize).Replace temp_dsp.Select(STPos, 1, BiasSize)
    Outwf_T3.Select(0, 1, BiasSize).Replace temp_dsp.Select(2 * STPos, 1, BiasSize)
    Outwf_T4.Select(0, 1, BiasSize).Replace temp_dsp.Select(3 * STPos, 1, BiasSize)
    
    Dim target_index As Long: target_index = 999
    Dim TempVal As Long
    Dim target_temp_value As Long
    
    ' Trimming first
    For i = 0 To Outwf_T1.SampleSize - 2
        If (Outwf_T1.Element(i) - 1024) >= 0 Then
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next
    
    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_temp_value = Outwf_T1.Subtract(1024).Abs.CalcMinimumValue(target_index)
    End If
'
    Outwf_T1.Element(Outwf_T1.SampleSize - 1) = target_index
    
    
    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T1.Element(Outwf_T1.SampleSize - 1) = TempVal
    
    For i = 0 To CLng((storeDSP1.SampleSize)) - 1
        storeDSP1.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            storeDSP1.Element(i) = 0
        End If
    Next i
    
    ' Trimming second
    target_index = 999
    For i = 0 To Outwf_T2.SampleSize - 2
        If (Outwf_T2.Element(i) - 1024) >= 0 Then
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next

    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_temp_value = Outwf_T2.Subtract(1024).Abs.CalcMinimumValue(target_index)
    End If
'
    Outwf_T2.Element(Outwf_T2.SampleSize - 1) = target_index

    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T2.Element(Outwf_T2.SampleSize - 1) = TempVal

    For i = 0 To CLng((StoreDSP2.SampleSize)) - 1
        StoreDSP2.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            StoreDSP2.Element(i) = 0
        End If
    Next i

    ' Trimming 3
    target_index = 999
    For i = 0 To Outwf_T3.SampleSize - 2
        If (Outwf_T3.Element(i) - 1024) >= 0 Then
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next

    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_temp_value = Outwf_T3.Subtract(1024).Abs.CalcMinimumValue(target_index)
    End If
'
    Outwf_T3.Element(Outwf_T3.SampleSize - 1) = target_index

    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T3.Element(Outwf_T3.SampleSize - 1) = TempVal

    For i = 0 To CLng((storeDSP3.SampleSize)) - 1
        storeDSP3.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            storeDSP3.Element(i) = 0
        End If
    Next i
  
    ' Trimming 4
    target_index = 999
    For i = 0 To Outwf_T4.SampleSize - 2
        If (Outwf_T4.Element(i) - 1024) >= 0 Then
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next

    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_temp_value = Outwf_T4.Subtract(1024).Abs.CalcMinimumValue(target_index)
    End If
'
    Outwf_T4.Element(Outwf_T4.SampleSize - 1) = target_index

    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T4.Element(Outwf_T4.SampleSize - 1) = TempVal

    For i = 0 To CLng((storeDSP3.SampleSize)) - 1
        storeDSP4.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            storeDSP4.Element(i) = 0
        End If
    Next i
    
End Function

Public Function Split_Dspwave_AUSPLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, ByRef Outwf_T1 As DSPWave, ByRef sda_measuredata As DSPWave, calc_data As DSPWave, Delta_Value As DSPWave, target_var As Double, ByVal calibration_target_value As Long, ByVal start_search As Long, ByVal search_width As Long, BinTarget As DSPWave, ByVal Interspersed_boolen As Boolean, ByVal z As Long, ByVal Lane_Sum As Long) As Long
                                                                                                                                                                                                                                                            
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim Interspersed_DSP As New DSPWave
    Dim Outwf_T1_Temp As New DSPWave
    Dim Count As Long
    current_loc = 0
    Count = 0
    
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    Outwf_T1_Temp.CreateConstant 0, sda_measuredata.Element(0) * Lane_Sum, DspLong
''    Outwf_T1.CreateConstant 0, width_Wf.SampleSize + 1, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        If width_Wf.Element(i) = search_width Then
            Outwf_T1_Temp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
        End If
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
                                                                                                                                                                                                                                                              
    If Interspersed_boolen = True Then
       Interspersed_DSP = Outwf_T1_Temp.Select(z, Lane_Sum, sda_measuredata.Element(0)).COPY
      Else
       Interspersed_DSP = Outwf_T1_Temp.Select(z * sda_measuredata.Element(0), 1, sda_measuredata.Element(0)).COPY
    End If
    
    For i = start_search To CLng(Interspersed_DSP.SampleSize) - 1
        calc_data.Element(i - start_search) = Interspersed_DSP.Element(i)
    Next i
    
''    Outwf_T1.Select(0, 1, OutWf.SampleSize).Replace OutWf.Select(0, 1, -1)
    
    ' 1.2 Calculate desired index
    Dim target_index As Long: target_index = 999
    Dim target_gap As Double: target_gap = 2013144
    For i = start_search To CLng((Interspersed_DSP.SampleSize)) - 1
        If (calc_data.Element(i - start_search) - calibration_target_value) >= 0 Then
            Delta_Value.Element(i) = Abs(calc_data.Element(i - start_search) - calibration_target_value)
            target_gap = Delta_Value.Element(i)
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next i
    
    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_gap = 2013144
        For i = start_search To CLng((Interspersed_DSP.SampleSize)) - 1
            Delta_Value.Element(i) = Abs(calc_data.Element(i - start_search) - calibration_target_value)
            If Delta_Value.Element(i) < target_gap Then
                target_gap = Delta_Value.Element(i)
                target_index = i
            End If
        Next
    End If
'
    'OutWf.Element(OutWf.SampleSize - 1) = target_index
    
    Dim TempVal As Long
    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    
''    Outwf_T1.Element(OutWf.SampleSize) = TempVal
    target_var = TempVal
    
    For i = 0 To CLng((BinTarget.SampleSize)) - 1
        BinTarget.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            BinTarget.Element(i) = 0
        End If
    Next i
                                                                                                                                                                                                                                                               
End Function

Public Function Split_Dspwave_AUSPLL_DESCENDING(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, ByRef Outwf_T1 As DSPWave, ByRef sda_measuredata As DSPWave, calc_data As DSPWave, Delta_Value As DSPWave, target_var As Double, ByVal calibration_target_value As Long, ByVal start_search As Long, ByVal search_width As Long, BinTarget As DSPWave, ByVal Interspersed_boolen As Boolean, ByVal z As Long, ByVal Lane_Sum As Long) As Long
                                                                                                                                                                                                                                                            
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim Interspersed_DSP As New DSPWave
    Dim Outwf_T1_Temp As New DSPWave
    Dim Count As Long
    current_loc = 0
    Count = 0
    
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    Outwf_T1_Temp.CreateConstant 0, sda_measuredata.Element(0) * Lane_Sum, DspLong
''    Outwf_T1.CreateConstant 0, width_Wf.SampleSize + 1, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
        If width_Wf.Element(i) = search_width Then
            Outwf_T1_Temp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
        End If
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
                                                                                                                                                                                                                                                              
    If Interspersed_boolen = True Then
       Interspersed_DSP = Outwf_T1_Temp.Select(z, Lane_Sum, sda_measuredata.Element(0)).COPY
      Else
       Interspersed_DSP = Outwf_T1_Temp.Select(z * sda_measuredata.Element(0), 1, sda_measuredata.Element(0)).COPY
    End If
    
    For i = start_search To CLng(Interspersed_DSP.SampleSize) - 1
        calc_data.Element(i - start_search) = Interspersed_DSP.Element(i)
    Next i
    
''    Outwf_T1.Select(0, 1, OutWf.SampleSize).Replace OutWf.Select(0, 1, -1)
    
    ' 1.2 Calculate desired index
    Dim target_index As Long: target_index = 999
    Dim target_gap As Double: target_gap = 2013144
    For i = start_search To CLng((Interspersed_DSP.SampleSize)) - 1
        If (calc_data.Element(i - start_search) - calibration_target_value) >= 0 Then
            Delta_Value.Element(i) = Abs(calc_data.Element(i - start_search) - calibration_target_value)
            target_gap = Delta_Value.Element(i)
            target_index = CLng((Interspersed_DSP.SampleSize)) - 1 - i      '@YM 20221102 IBIZA CIOAUSPLL Descending Order Capture
            Exit For  'once larger then use it for 1st read
        End If
    Next i
    
    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_gap = 2013144
        For i = start_search To CLng((Interspersed_DSP.SampleSize)) - 1
            Delta_Value.Element(i) = Abs(calc_data.Element(i - start_search) - calibration_target_value)
            If Delta_Value.Element(i) < target_gap Then
                target_gap = Delta_Value.Element(i)
                target_index = CLng((Interspersed_DSP.SampleSize)) - 1 - i      '@YM 20221102 IBIZA CIOAUSPLL Descending Order Capture
            End If
        Next
    End If
'
    'OutWf.Element(OutWf.SampleSize - 1) = target_index
    
    Dim TempVal As Long
    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    
''    Outwf_T1.Element(OutWf.SampleSize) = TempVal
    target_var = TempVal
    
    For i = 0 To CLng((BinTarget.SampleSize)) - 1
        BinTarget.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            BinTarget.Element(i) = 0
        End If
    Next i
                                                                                                                                                                                                                                                               
End Function

Public Function DSPWF_AVG(ByVal InWf As DSPWave, ByVal DSPWave_Dictionry_amount As DSPWave, ByVal DSPWave_Split_Bit_perDic As DSPWave, _
                        ByVal DSPWave_2SComplement_Flag As DSPWave, ByVal DSPWave_LSBtoMSB_Flag As DSPWave, ByRef DSPWave_AvgDecBySet As DSPWave, ByRef DSPWave_AvgDec2BinBySet As DSPWave, ByRef DSPWave_STDEV As DSPWave) As Long
    'Update for IVDM function TTR (DSP process from local to DSPPC )-- 20220105
    Dim DSPWave_Dec As New DSPWave
    Dim DSPWave_SumDec As New DSPWave
    Dim DSPWave_SumDecBySet As New DSPWave
    
    Dim SUM_Temp As New DSPWave
    Dim AVG_Temp As New DSPWave
    
    Dim Denominator As Long
    Dim LSBtoMSB_Flag As Long
    Dim TwoComplement_Flag As Long
    Dim i As Long
    Dim j As Long
    Dim Split_bit As Long: Split_bit = 0
    Dim Max_bit As Long: Max_bit = 0
    Dim Max_set As Long: Max_set = 0
    Dim pIndex As Long: pIndex = 0
    
    Dim InWf_Par As New DSPWave
    Dim InWf_Par_MSB1st As New DSPWave
    DSPWave_SumDec.CreateConstant 0, 1, DspLong
    SUM_Temp.CreateConstant 0, 1, DspLong
    AVG_Temp.CreateConstant 0, 1, DspDouble
    
    Max_set = DSPWave_Dictionry_amount.SampleSize
    DSPWave_SumDecBySet.CreateConstant 0, Max_set, DspLong
    DSPWave_AvgDecBySet.CreateConstant 0, Max_set, DspDouble
    
    Dim DSPWave_ALL_DATA As New DSPWave                 'New for IVDM STDEV -- 20230116
    Dim dSTDEV As Double
    Dim dMean As Double
    
    DSPWave_STDEV.CreateConstant 0, 1, DspDouble        'New for IVDM STDEV -- 20230116
    
    For j = 0 To DSPWave_Dictionry_amount.SampleSize - 1
        InWf_Par.CreateConstant 0, DSPWave_Split_Bit_perDic.Element(j), DspLong
        InWf_Par_MSB1st.CreateConstant 0, DSPWave_Split_Bit_perDic.Element(j), DspLong
        DSPWave_Dec.CreateConstant 0, 1, DspLong
        
        Max_bit = DSPWave_Split_Bit_perDic.Element(j)
        DSPWave_ALL_DATA.CreateConstant 0, InWf.SampleSize / Max_bit, DspDouble 'New for IVDM STDEV -- 20230116
        
        Denominator = DSPWave_Dictionry_amount.Element(j)
        TwoComplement_Flag = DSPWave_2SComplement_Flag.Element(j)
        LSBtoMSB_Flag = DSPWave_LSBtoMSB_Flag.Element(j)

        For i = pIndex To InWf.SampleSize - 1 Step Max_bit
            InWf_Par = InWf.Select(i, , Max_bit).COPY
            If LSBtoMSB_Flag = 1 Then
                Call Split_Dspwave_Reverse(InWf_Par, InWf_Par_MSB1st)
                InWf_Par = InWf_Par_MSB1st.COPY
            End If
            
            If TwoComplement_Flag = 0 Then
                DSPWave_Dec = InWf_Par.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, InWf_Par.SampleSize, 0, Bit0IsMsb)
            Else
                Call DSP_2S_Complement_To_SignDec(InWf_Par, DSPWave_Split_Bit_perDic.Element(j), DSPWave_Dec)
            End If
            
            DSPWave_ALL_DATA.Element(i / Max_bit) = DSPWave_Dec.Element(0)     'New for IVDM STDEV -- 20230116
        Next i
    Next j
    
    dMean = DSPWave_ALL_DATA.CalcMeanWithStdDev(dSTDEV)   'New for IVDM STDEV -- 20230116
    
    DSPWave_AvgDecBySet.Element(0) = dMean
    DSPWave_STDEV.Element(0) = dSTDEV                     'New for IVDM STDEV -- 20230116
    
    Call DSPWaveDecToBinary(DSPWave_AvgDecBySet.ConvertDataTypeTo(DspLong), Max_bit, DSPWave_AvgDec2BinBySet)
                    
End Function


Public Function DSPWF_OFFSET(ByVal DSPWave_LowBin As DSPWave, ByVal lVoltage_Low As Double, ByVal DSPWave_HighBin As DSPWave, ByVal lVoltage_High As Double, _
                          ByVal lYintercept_spec As Double, ByRef DSPWave_SlopeDec As DSPWave, ByRef DSPWave_YaxisBDec As DSPWave, _
                          ByRef DSPWave_InterpolateDec As DSPWave, ByRef DSPWave_OffsetDec As DSPWave, ByRef DSPWave_OffsetDec2BinBySet As DSPWave, ByVal Flag_2sComplement As Boolean) As Long
    'Update for IVDM function TTR (DSP process from local to DSPPC )-- 20220105
    Dim DSPWave_LowDec As New DSPWave
    Dim DSPWave_HighDec As New DSPWave
    Dim SL_BitWidth As Long
    Dim DSPWave_OffsetDec_Rev As New DSPWave            'Upate for negative offset value -- 20220629
    
    DSPWave_LowDec.CreateConstant 0, 1, DspDouble
    DSPWave_HighDec.CreateConstant 0, 1, DspDouble
    DSPWave_OffsetDec_Rev.CreateConstant 0, 1, DspLong
    
    Call BinToDec(DSPWave_LowBin, DSPWave_LowDec)
    Call BinToDec(DSPWave_HighBin, DSPWave_HighDec)

    SL_BitWidth = DSPWave_LowBin.SampleSize
    DSPWave_SlopeDec = DSPWave_HighDec.Subtract(DSPWave_LowDec).divide(lVoltage_High - lVoltage_Low)
    DSPWave_YaxisBDec = DSPWave_LowDec.Subtract(DSPWave_SlopeDec.Multiply(lVoltage_Low))
    DSPWave_InterpolateDec = DSPWave_YaxisBDec.Add(DSPWave_SlopeDec.Multiply(0.95))
    DSPWave_OffsetDec = DSPWave_InterpolateDec.Negate.Add(lYintercept_spec)
    If Flag_2sComplement = True Then
        DSPWave_OffsetDec_Rev = DSPWave_OffsetDec
    Else
        DSPWave_OffsetDec_Rev = DSPWave_OffsetDec.ConvertDataTypeTo(DspLong)
        If DSPWave_OffsetDec_Rev.Element(0) < 0 Then
            DSPWave_OffsetDec_Rev = DSPWave_OffsetDec_Rev.Subtract(128).Abs         'Upate for negative offset value -- 20220704
        End If
    End If

    Call DSPWaveDecToBinary(DSPWave_OffsetDec_Rev.ConvertDataTypeTo(DspLong), SL_BitWidth, DSPWave_OffsetDec2BinBySet)
    
    
End Function

Public Function Split_Dspwave_Partial_Index(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, IdexWf As DSPWave, ByVal Bit0Force0 As Boolean) As Long
    'Update for MRR function -- 20230206

    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim TempWf As New DSPWave
    '' current_loc = 0 + offset ''additional_loc = 16
    '' 20170705 - Solve for over 32 bits to decimal
    'OutWf.CreateConstant 0, width_Wf.SampleSize, DspLong
    For i = 0 To (width_Wf.SampleSize - 1)
        dec_wave = InWf.Select(IdexWf.Element(i), , width_Wf.Element(i)).COPY
        If Bit0Force0 = True Then
            If width_Wf.Element(i) = 8 Then         'Bit width = 8bits
                dec_wave.Element(7) = 0           'Force LSB bit as 0 for ignore MRR8 LSB bit -- 20230207
            End If
            If width_Wf.Element(i) = 16 Then        'Bit width = 16bits
                dec_wave.Element(7) = 0
                dec_wave.Element(15) = 0          'Force LSB bit as 0 for ignore MRR8 LSB bit -- 20230207
            End If
        End If

        If i = 0 Then
            TempWf = dec_wave.COPY
        Else
            TempWf = TempWf.Concatenate(dec_wave).COPY
        End If
    Next i
    OutWf = TempWf.COPY
    
End Function

Public Function post_process_capwave_variance(capWave As DSPWave, ByRef Avg As Double, ByRef stdDev As Double, ByRef maxmin As Double) As Long

    If capWave.SampleSize = 1 Then
        Avg = capWave.CalcMean
        stdDev = 0
    Else
        Avg = capWave.CalcMeanWithStdDev(stdDev)
    End If
    maxmin = capWave.CalcMaximumValue - capWave.CalcMinimumValue
End Function

'[20230915][All][Neil] DRAM MRR New feature implement
Public Function Split_Dspwave_Partial_Index_New(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, IdexWf As DSPWave, ByVal BitForce0_Num As Long) As Long
    'Update for MRR function -- 20230206
    'Add option to fill floating bits amount
    
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim TempWf As New DSPWave
    Dim bit_index As Long
    Dim BitForce0_index As Long
    For i = 0 To (width_Wf.SampleSize - 1)
        dec_wave = InWf.Select(IdexWf.Element(i), , width_Wf.Element(i)).COPY
        
        If BitForce0_Num <> 0 Then
        
            For bit_index = 8 To dec_wave.SampleSize - 1 Step 8
                For BitForce0_index = 1 To BitForce0_Num
                    dec_wave.Element(bit_index - BitForce0_index) = 0
                Next BitForce0_index
            Next bit_index
        End If
        
        If i = 0 Then
            TempWf = dec_wave.COPY
        Else
            TempWf = TempWf.Concatenate(dec_wave).COPY
        End If
    Next i
    OutWf = TempWf.COPY
    
End Function


'[20231107][All][Neil] Duty Cycle Test Methodology on Digital and DCVI instrument
Public Function DSP_Measure_Jitter(ByVal in_DspWave As DSPWave, ByRef Out_DutyCycle As Double) As Long
    ' Use for Measure DutyCycle for Digital instrument
    Dim calcDspWave As New DSPWave
    Dim SerialDspWave As New DSPWave
    Dim RJ As Double
    Dim DDJ As Double
    Dim MeasuredUI As Double
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim period As Double
    Dim Status As Long
    Dim index As Long
    calcDspWave = in_DspWave.COPY
    SerialDspWave = calcDspWave.ConvertStreamTo(tldspSerial, 32, 0, Bit0IsLsb)
    Out_DutyCycle = SerialDspWave.CalcMean
End Function

'[20231107][All][Neil] Duty Cycle Test Methodology on Digital and DCVI instrument
Public Function DSP_CALC_FREQ_DUTY_FROM_SPECTRUM(CAPMEM_DATA As DSPWave, ByVal SampleRate As Double, ByVal SampleSize As Double, ByVal FirstOneCycleFetch As Boolean, ByVal AVG_Times As Long, Return_Frequency As DSPWave, Return_DutyCycle As DSPWave) As Long
    'Format : DSP_CALC_FREQ_DUTY_FROM_SPECTRUM(CAPMEM_Data,SampleRate,SampleSize,ThresholdVoltage,FirstOneCycleFetch,AVG_Times,Freq_Result,DutyResult)
    Dim DSPWF_SpectrumData As New DSPWave
    Dim DSPWF_OneCycleRawData As New DSPWave
    Dim DSPWF_Hi_Pulse_Samples As New DSPWave
    Dim SL_SamplePoint As Long
    Return_Frequency.CreateConstant 0, 1, DspDouble
    Return_DutyCycle.CreateConstant 0, 1, DspDouble
    DSPWF_Hi_Pulse_Samples.CreateConstant 0, 1, DspDouble
    Dim V_Avg As Double
    
    'Meas Frequency
    Set DSPWF_SpectrumData = CAPMEM_DATA.Spectrum
    Return_Frequency.Element(0) = DSPWF_SpectrumData.CalcFrequencyFromSpectrum

    'Meas Duty
    SL_SamplePoint = Int(SampleRate / Return_Frequency.Element(0)) 'SamplePoint=Cap_SampleRate/Meas_Freq [ Freq Domain]
    If SL_SamplePoint * AVG_Times > CAPMEM_DATA.SampleSize Then SL_SamplePoint = CAPMEM_DATA.SampleSize / AVG_Times
    DSPWF_OneCycleRawData.CreateConstant 0, SL_SamplePoint, DspLong
    
    
    If FirstOneCycleFetch = True Then
        DSPWF_OneCycleRawData = CAPMEM_DATA.Select(0, , SL_SamplePoint * AVG_Times).COPY ' Fetch first cycle data from CAP memory
    Else
        DSPWF_OneCycleRawData = CAPMEM_DATA.Select(CAPMEM_DATA.SampleSize - (SL_SamplePoint * AVG_Times), , (SL_SamplePoint * AVG_Times)).COPY ' Fetch first cycle data from CAP memory
    End If
    ' Calculate Duty Cycle
    V_Avg = DSPWF_OneCycleRawData.CalcMinimumValue + 0.5 * (DSPWF_OneCycleRawData.CalcMaximumValue - DSPWF_OneCycleRawData.CalcMinimumValue)
    DSPWF_Hi_Pulse_Samples.Element(0) = DSPWF_OneCycleRawData.FindIndices(GreaterThanOrEqualTo, V_Avg).SampleSize 'Get Hi-State samples
    Return_DutyCycle = DSPWF_Hi_Pulse_Samples.divide(SL_SamplePoint * AVG_Times)  'DutyCycle = Hi-Samples/Fetch_OneCycleSamples
End Function

Public Function Split_Dspwave_PCIE5PLL_DESCENDING(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, target_var As Double, _
ByVal calibration_target_value_low As Long, ByVal calibration_target_value_high As Long, _
ByVal start_search As Long, Bin_Target_high As DSPWave, ByVal Interspersed_boolen As Boolean, _
ByVal z As Long, ByVal Lane_Sum As Long, ByVal Each_lane_size As Long, ByVal offset As Long) As Long

'20240131 Hidra by YM

    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim Interspersed_DSP As New DSPWave
    Dim target_var_low As Long
    Dim target_var_high As Long
    
    current_loc = 0
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
            dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).COPY
            OutWf.Element(i) = dec_wave.Element(0)
            current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    
    If Interspersed_boolen = True Then
        Interspersed_DSP = OutWf.Select(z, Lane_Sum, Each_lane_size).COPY
    Else
        Interspersed_DSP = OutWf.Select(z * Each_lane_size, 1, Each_lane_size).COPY
    End If
     
    For i = start_search To CLng(Interspersed_DSP.SampleSize) - 1
        calc_data.Element(i - start_search) = Interspersed_DSP.Element(i)
    Next i
    
    Dim j As Integer
    Dim k As Integer

    For i = start_search To CLng((Interspersed_DSP.SampleSize)) - 1
        'Find First Zero index
        If Interspersed_DSP.Element(i) = calibration_target_value_low Then
            target_var_low = i
            Exit For
        End If
    Next i
    
    target_var_high = 9999
    
    For j = target_var_low + 1 To CLng((Interspersed_DSP.SampleSize)) - 1
        'Find Second index
        If Interspersed_DSP.Element(j) >= calibration_target_value_high Then
            target_var_high = CLng((Interspersed_DSP.SampleSize)) - 1 - j   'Decending Target = Total Size - j
            Exit For
        End If
    Next j
    
    If target_var_high = 9999 Then
        target_var_high = 0
    End If
    
    'target_var_low = target_var_low + offset
    target_var_high = target_var_high + offset
    ''''''''''''' Decimal to Binary dspwave for dictionary'''''''''''''''
    
    Dim TempVal As Long
    
    TempVal = target_var_high
    target_var = TempVal
    
    For i = 0 To CLng((Bin_Target_high.SampleSize)) - 1
        Bin_Target_high.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_high.Element(i) = 0
        End If
    Next i
    
End Function

Public Function DSP_SelectBits(ByVal InWf As DSPWave, ByRef OutWf As DSPWave, ByVal start_bit As Long, ByVal length As Long) As Long
  
    OutWf = InWf.Select(start_bit, 1, length).COPY

End Function
Public Function DSPWf_Concatenate_MRR5678(ByRef OutputDspWave As DSPWave, ByVal InDSPwave As DSPWave, ByVal InDSPwave1 As DSPWave, ByVal InDSPwave2 As DSPWave, ByVal InDSPwave3 As DSPWave) As Long
    
    InDSPwave = InDSPwave.ConvertDataTypeTo(DspLong)
    
    OutputDspWave = InDSPwave.COPY
    
    InDSPwave1 = InDSPwave1.ConvertDataTypeTo(DspLong)
    
    OutputDspWave = OutputDspWave.Concatenate(InDSPwave1)
    
    InDSPwave2 = InDSPwave2.ConvertDataTypeTo(DspLong)
    
    OutputDspWave = OutputDspWave.Concatenate(InDSPwave2)

    InDSPwave3 = InDSPwave3.ConvertDataTypeTo(DspLong)
    
    OutputDspWave = OutputDspWave.Concatenate(InDSPwave3)
    
End Function


Public Function DSP_DigCap_Frequency_Transfer(ByVal InWf As DSPWave, ByVal ratio As Double, ByRef OutWf As DSPWave) As Long
    
    OutWf = InWf.Multiply(ratio).Multiply(1000000)

End Function
