Attribute VB_Name = "DSP_BACK"
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


Public Function pulseMeas(ByVal wave As DSPWave, PWHigh As Double, PWLow As Double, Period As Double, Status As Long) As Long
        

        
    Call wave.measuretime(PWHigh, PWLow, Period, Status)
    
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
        calcDspWave = CapDSPWave.Copy
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
        SegWf = InWf.Select(SegIdx * SegWidth, 1, SegWidth).Copy
        If MdllWrdWidth <> 0 Then
            tmpwf = SegWf.Select(0, 1, MdllWrdWidth).Copy
            MdllWf.Element(SegIdx) = tmpwf.ConvertStreamTo(tldspParallel, MdllWrdWidth, 0, Bit0IsMsb).Element(0)
            tmpwf.Clear
        End If
        
        For BistIdx = 0 To BistsPerMdll - 1
            Dim st0 As Long, st1 As Long
            st0 = EyeStrobes * BistIdx
            st1 = SegIdx * (SegWidth - MdllWrdWidth) + st0
            tmpwf = SegWf.Select(MdllWrdWidth + st0, 1, EyeStrobes).Copy
            EyeWf.Select(st1, 1, EyeStrobes).Replace tmpwf
            tmpwf.Clear
        Next BistIdx
    Next SegIdx
    
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

    DataWf_Binary = InWf.Select(StartIndex, , WrdWdth).Copy
    DataWf = InWf.Select(StartIndex, , WrdWdth).ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

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
                DigSrcWf = SingleWordWf.Copy
            End If
        Next DataIdx
        DigSrcWf = DigSrcWf.Repeat(NoOfRepeats)
    End If
    
    
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
    TempDSP = InWf.Select(StartBit, 1, BitLength).Copy
    
    DataWf = TempDSP.ConvertStreamTo(tldspParallel, BitLength, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
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
            DSP_Input_UpperBIN = DSP_Input_Update.Select(i, , length).Copy
            DSP_Input_BelowBIN = DSP_Input_Update.Select(i + length, , length).Copy
        Else
            DSP_Input_UpperBIN = DSP_Input_UpperBIN.Concatenate(DSP_Input_Update.Select(i, , length).Copy)
            DSP_Input_BelowBIN = DSP_Input_BelowBIN.Concatenate(DSP_Input_Update.Select(i + length, , length).Copy)
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
            DSP_Input_UpperBIN_1 = DSP_Input_Update.Select(i, , length).Copy
            DSP_Input_UpperBIN_2 = DSP_Input_Update.Select(i + length, , length).Copy
            DSP_Input_BelowBIN_1 = DSP_Input_Update.Select(i + length * 2, , length).Copy
            DSP_Input_BelowBIN_2 = DSP_Input_Update.Select(i + length * 3, , length).Copy
        Else
            DSP_Input_UpperBIN_1 = DSP_Input_UpperBIN_1.Concatenate(DSP_Input_Update.Select(i, , length).Copy)
            DSP_Input_UpperBIN_2 = DSP_Input_UpperBIN_2.Concatenate(DSP_Input_Update.Select(i + length, , length).Copy)
            DSP_Input_BelowBIN_1 = DSP_Input_BelowBIN_1.Concatenate(DSP_Input_Update.Select(i + length * 2, , length).Copy)
            DSP_Input_BelowBIN_2 = DSP_Input_BelowBIN_2.Concatenate(DSP_Input_Update.Select(i + length * 3, , length).Copy)
        
        End If
    Next i
            
    DSP_Input_UpperDEC_1 = DSP_Input_UpperBIN_1.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_UpperDEC_2 = DSP_Input_UpperBIN_2.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_BelowDEC_1 = DSP_Input_BelowBIN_1.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_BelowDEC_2 = DSP_Input_BelowBIN_2.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    
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

Public Function Split_Dspwave_CIOCALC(Outwf_T1 As DSPWave, V2 As Double, V3 As Double, V4 As Double, V5 As Double, V6 As Double, V7 As Double, V8 As Double, V9 As Double, V10 As Double, V11 As Double, V12 As Double, V13 As Double, V14 As Double, V15 As Double, V16 As Double, V17 As Double, V18 As Double, V19 As Double, V20 As Double, V21 As Double, V22 As Double, V23 As Double, _
                                V24 As Double, V25 As Double, storeDSP As DSPWave) As Long
                                                                                                                                                                                                                                                               
                                                                                                                                                                                                                                                               
    ' Special calculation for T1
    ' 1.1 store measured data
    Dim d_temp(24) As Double
    d_temp(0) = V2
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
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).Copy
        If width_Wf.Element(i) = 16 Then
            'dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
            temp_dsp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
            'current_loc = current_loc + width_Wf.Element(i)
        End If
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    Outwf_T1 = temp_dsp.Select(1, 1, 25).Copy
                                                                                                                                                                                                                                                               
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
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).Copy
        If width_Wf.Element(i) = 16 Then
            temp_dsp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
        End If

        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    Outwf_T1 = temp_dsp.Select(1, 1, Count / 2 - 2).Copy
    Outwf_T2 = temp_dsp.Select(Count / 2 + 1, 1, Count / 2 - 2).Copy
    
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
        DSP_SignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        Call DSP_GrayCode2Bin(False, DSP_SignedGray, Out_Wf_Binary, dec_wave)
        If Index_SignedGray <> DSPSignedGray_StartBit.SampleSize - 1 Then
            Index_SignedGray = Index_SignedGray + 1
        End If
    ElseIf current_loc = DSPUnSignedGray_StartBit.Element(Index_UnSignedGray) Then
        DSP_UnSignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
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
        DSP_UnSignedBinary = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        dec_wave = DSP_UnSignedBinary.ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
    End If
    If i = 0 Then
        OutWf = dec_wave.Copy
    Else
        OutWf = OutWf.Concatenate(dec_wave)
    End If
    current_loc = current_loc + width_Wf.Element(i)
Next i
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
    DataWf = InWf.Copy

    For i = 0 To AssignmentDSPWave.SampleSize
        If AssignmentDSPWave.Element(i) = 0 Then
            SelsetOffset = i * CalculateSize
            Exit For
        Else
            SelsetOffset = 0
        End If
    Next i


    CalculateDSP = DataWf.Select(SelsetOffset, 1, CalculateSize).Copy
    DataWf = CalculateDSP.Copy
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
    CalculateDSP = DataWf.Copy
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

Public Function MergeWave(InPutWaveA As DSPWave, InPutWaveB As DSPWave, Outputwave As DSPWave) As Long

    Dim TempWave As New DSPWave
    
    TempWave = InPutWaveA.Concatenate(InPutWaveB)
    Outputwave = TempWave.Copy

End Function
  

Public Function NANDIO_DigCap(InDSPWave As DSPWave, PVTPNR_int As Long, OutDspWave As DSPWave, WaveStr As Double) As Long  '(CUS_Str_DigCapData As String, InDspWave As DSPWave, PVTPNR_Int As Long, OutDSPWave As DSPWave) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
    
    
    ''===========InWaveStr=============
''        For i = 0 To InDspWave.SampleSize - 1
''            AA = AA * 10 + InDspWave.Element(i)
''        Next i
    
    ''================================PVTP============================================
    'If InStr(UCase(CUS_Str_DigCapData), "PVTP") <> 0 Then
        
'''''''        Index = InDspWave.CalcSum
'''''''        Call InDspWave.CalcMaximumValue(VIndex)
'''''''        PVTPNR_int = 31 - VIndex
'''''''
'''''''        If 32 - VIndex - Index <> 0 Then
'''''''             For i = 0 To InDspWave.SampleSize - 1
'''''''
'''''''                If InDspWave.Element(i) = 1 Then
'''''''                    Index = i
'''''''                    Exit For
'''''''                End If
'''''''
'''''''
'''''''             Next i
'''''''             PVTPNR_int = 31 - Index
'''''''        End If
'''''''        VIndex = PVTPNR_int
'''''''        For i = 0 To 5
'''''''            If PVTPNR_int > 0 Then
'''''''                OutDSPWave.Element(i) = PVTPNR_int Mod 2
'''''''                PVTPNR_int = Int(PVTPNR_int / 2)
'''''''
'''''''            Else
'''''''                OutDSPWave.Element(i) = 0
'''''''
'''''''            End If
'''''''
'''''''        Next i
'''''''        PVTPNR_int = VIndex
        
    'End If

    ''================================PVTN============================================
    'If InStr(UCase(CUS_Str_DigCapData), "PVTN") <> 0 Then
'''        Index = InDspWave.CalcSum
'''        Call InDspWave.CalcMaximumValue(VIndex)
'''        PVTPNR_int = VIndex
'''
'''        If 32 - PVTPNR_int - Index <> 0 Then
'''             For i = 0 To InDspWave.SampleSize - 1
'''
'''                If InDspWave.Element(i) = 1 Then
'''                    Index = i
'''                    Exit For
'''                End If
'''                PVTPNR_int = Index
'''
'''             Next i
'''
'''        End If
'''
'''        VIndex = PVTPNR_int
'''        For i = 0 To 5
'''            If PVTPNR_int > 0 Then
'''                OutDSPWave.Element(i) = PVTPNR_int Mod 2
'''                PVTPNR_int = Int(PVTPNR_int / 2)
'''            Else
'''                OutDSPWave.Element(i) = 0
'''            End If
'''
'''        Next i
'''        PVTPNR_int = VIndex
        
    'End If

    ''================================PVTR============================================
    'If InStr(UCase(CUS_Str_DigCapData), "PVTR") <> 0 Then
        index = InDSPWave.CalcSum
        Call InDSPWave.CalcMinimumValue(vIndex)
        PVTPNR_int = vIndex
        
        If 32 - PVTPNR_int - index <> 0 Then
             For i = 0 To InDSPWave.SampleSize - 1
                
                If InDSPWave.Element(i) = 1 Then
                    index = i
                    Exit For
                End If
                PVTPNR_int = index
                
             Next i
        
        End If
        
        vIndex = PVTPNR_int
        For i = 14 To 0 Step -3
            If PVTPNR_int > 0 Then
                OutDspWave.Element(i) = PVTPNR_int Mod 2
                OutDspWave.Element(i - 1) = PVTPNR_int Mod 2
                OutDspWave.Element(i - 2) = PVTPNR_int Mod 2
                PVTPNR_int = Int(PVTPNR_int / 2)
            Else
                OutDspWave.Element(i) = 0
                OutDspWave.Element(i - 1) = 0
                OutDspWave.Element(i - 2) = 0
            End If
        Next i
        PVTPNR_int = vIndex
        
    'End If
 If False Then
    ''===========================CAL_UP_DOWN============================================
    'If InStr(UCase(CUS_Str_DigCapData), "CAL_UP_DOWN") <> 0 Then
            
        OutDspWave = InDSPWave.Select(0, 7, 10).Copy
            
            
'''            For j = 0 To 9
'''                out_DSPWave_tmp(0).CreateConstant 0, 63 '''pattern change from 63 to 64
'''                show_str_PVT_CAL_UP = ""
'''                show_str_PVT_CAL_DOWN = ""
'''                show_obsIO = ""
'''
'''                For i = j To 629 Step 10 '''pattern change from 63 to 64
'''                    If i <> 0 Then
'''                        out_DSPWave_tmp(0).Element(Int(i / 10)) = OutDSPWave(DSPWaveIndex)(Site).Element(i)
'''                        show_obsIO = show_obsIO & OutDSPWave(DSPWaveIndex)(Site).Element(i)
'''                    Else
'''                        out_DSPWave_tmp(0).Element(0) = OutDSPWave(DSPWaveIndex)(Site).Element(i)
'''                        show_obsIO = OutDSPWave(DSPWaveIndex)(Site).Element(i)
'''                    End If
'''                Next i
'''
'''                Call Search_Tran_Bit(out_DSPWave_tmp, 63, Site, Tran_PVTPNR, Sum_CAL_UP_DOWN(j), CUS_Str_MainProgram)  '''pattern change from 63 to 64
'''                Tran_CAL_UP_DOWN(j) = Tran_PVTPNR
'''                If Tran_PVTPNR <> 999 Then
'''                    G_obsIO_up(j).CreateConstant 0, 6
'''                    G_obsIO_down(j).CreateConstant 0, 6
'''
'''                    For i = 0 To 5
'''                        If Tran_PVTPNR > 0 Then
'''                            G_obsIO_up(j)(Site).Element(i) = Tran_PVTPNR Mod 2
'''                            Tran_PVTPNR = Int(Tran_PVTPNR / 2)
'''                            show_str_PVT_CAL_UP = show_str_PVT_CAL_UP & G_obsIO_up(j)(Site).Element(i)
'''                        Else
'''                            G_obsIO_up(j)(Site).Element(i) = 0
'''                            show_str_PVT_CAL_UP = show_str_PVT_CAL_UP & G_obsIO_up(j)(Site).Element(i)
'''                        End If
'''                    Next i
'''
'''                    Tran_PVTPNR = 62 - Tran_CAL_UP_DOWN(j)
'''                    For i = 0 To 5
'''                        If Tran_PVTPNR > 0 Then
'''                            G_obsIO_down(j)(Site).Element(i) = Tran_PVTPNR Mod 2
'''                            Tran_PVTPNR = Int(Tran_PVTPNR / 2)
'''                            show_str_PVT_CAL_DOWN = show_str_PVT_CAL_DOWN & G_obsIO_down(j)(Site).Element(i)
'''                        Else
'''                            G_obsIO_down(j)(Site).Element(i) = 0
'''                            show_str_PVT_CAL_DOWN = show_str_PVT_CAL_DOWN & G_obsIO_down(j)(Site).Element(i)
'''                        End If
'''                    Next i
'''
'''                End If
'''            Next j
    
    'End If
    
 End If
End Function

Public Function NANDIO_CAL_UD_Wave(OutDSPWave0 As DSPWave, OutDSPWave1 As DSPWave, OutDSPWave2 As DSPWave, OutDSPWave3 As DSPWave, OutDSPWave4 As DSPWave, OutDSPWave5 As DSPWave, OutDSPWave6 As DSPWave, OutDSPWave7 As DSPWave, OutDSPWave8 As DSPWave, OutDSPWave9 As DSPWave, _
                                TNSPNT0 As Long, TNSPNT1 As Long, TNSPNT2 As Long, TNSPNT3 As Long, TNSPNT4 As Long, TNSPNT5 As Long, TNSPNT6 As Long, TNSPNT7 As Long, TNSPNT8 As Long, TNSPNT9 As Long) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
    Dim j As Integer
    Dim OutDspWave(9) As New DSPWave
    Dim TNSPNT(9) As Long
    
    TNSPNT(0) = TNSPNT0
    TNSPNT(1) = TNSPNT1
    TNSPNT(2) = TNSPNT2
    TNSPNT(3) = TNSPNT3
    TNSPNT(4) = TNSPNT4
    TNSPNT(5) = TNSPNT5
    TNSPNT(6) = TNSPNT6
    TNSPNT(7) = TNSPNT7
    TNSPNT(8) = TNSPNT8
    TNSPNT(9) = TNSPNT9
    
    For i = 0 To 9
        OutDspWave(i).CreateConstant 0, 12
        vIndex = TNSPNT(i)
        For j = 0 To 5
            If vIndex > 0 Then
                OutDspWave(i).Element(j) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDspWave(i).Element(j) = 0
                
            End If
    
        Next j
        
        vIndex = 62 - TNSPNT(i)
        For j = 6 To 11
            If vIndex > 0 Then
                OutDspWave(i).Element(j) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDspWave(i).Element(j) = 0
                
            End If
    
        Next j
    Next i
        
    OutDSPWave0 = OutDspWave(0).Copy
    OutDSPWave1 = OutDspWave(1).Copy
    OutDSPWave2 = OutDspWave(2).Copy
    OutDSPWave3 = OutDspWave(3).Copy
    OutDSPWave4 = OutDspWave(4).Copy
    OutDSPWave5 = OutDspWave(5).Copy
    OutDSPWave6 = OutDspWave(6).Copy
    OutDSPWave7 = OutDspWave(7).Copy
    OutDSPWave8 = OutDspWave(8).Copy
    OutDSPWave9 = OutDspWave(9).Copy
    
End Function




Public Function NANDIO_CAL_DOWN_Wave(InDSPWave As DSPWave, OutDSPWave0 As DSPWave, OutDSPWave1 As DSPWave, OutDSPWave2 As DSPWave, OutDSPWave3 As DSPWave, OutDSPWave4 As DSPWave, OutDSPWave5 As DSPWave, OutDSPWave6 As DSPWave, OutDSPWave7 As DSPWave, OutDSPWave8 As DSPWave, OutDSPWave9 As DSPWave, _
                                TNSPNT0 As Long, TNSPNT1 As Long, TNSPNT2 As Long, TNSPNT3 As Long, TNSPNT4 As Long, TNSPNT5 As Long, TNSPNT6 As Long, TNSPNT7 As Long, TNSPNT8 As Long, TNSPNT9 As Long) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
    Dim j As Integer
    Dim OutDspWave(9) As New DSPWave
    Dim TNSPNT(9) As Long
    
    
    TNSPNT(0) = TNSPNT0
    TNSPNT(1) = TNSPNT1
    TNSPNT(2) = TNSPNT2
    TNSPNT(3) = TNSPNT3
    TNSPNT(4) = TNSPNT4
    TNSPNT(5) = TNSPNT5
    TNSPNT(6) = TNSPNT6
    TNSPNT(7) = TNSPNT7
    TNSPNT(8) = TNSPNT8
    TNSPNT(9) = TNSPNT9
    
    For i = 0 To 9
        OutDspWave(i).CreateConstant 0, 6
        vIndex = 62 - TNSPNT(i)
        For j = 0 To 5
            If vIndex > 0 Then
                OutDspWave(i).Element(j) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDspWave(i).Element(j) = 0
                
            End If
    
        Next j
    Next i
        
    OutDSPWave0 = OutDspWave(0).Copy
    OutDSPWave1 = OutDspWave(1).Copy
    OutDSPWave2 = OutDspWave(2).Copy
    OutDSPWave3 = OutDspWave(3).Copy
    OutDSPWave4 = OutDspWave(4).Copy
    OutDSPWave5 = OutDspWave(5).Copy
    OutDSPWave6 = OutDspWave(6).Copy
    OutDSPWave7 = OutDspWave(7).Copy
    OutDSPWave8 = OutDspWave(8).Copy
    OutDSPWave9 = OutDspWave(9).Copy
    
    

End Function

Public Function NANDIO_CAL_UP_Wave(InDSPWave As DSPWave, OutDSPWave0 As DSPWave, OutDSPWave1 As DSPWave, OutDSPWave2 As DSPWave, OutDSPWave3 As DSPWave, OutDSPWave4 As DSPWave, OutDSPWave5 As DSPWave, OutDSPWave6 As DSPWave, OutDSPWave7 As DSPWave, OutDSPWave8 As DSPWave, OutDSPWave9 As DSPWave, _
                                TNSPNT0 As Long, TNSPNT1 As Long, TNSPNT2 As Long, TNSPNT3 As Long, TNSPNT4 As Long, TNSPNT5 As Long, TNSPNT6 As Long, TNSPNT7 As Long, TNSPNT8 As Long, TNSPNT9 As Long) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
    Dim j As Integer
    Dim OutDspWave(9) As New DSPWave
    Dim TNSPNT(9) As Long
    
    
    TNSPNT(0) = TNSPNT0
    TNSPNT(1) = TNSPNT1
    TNSPNT(2) = TNSPNT2
    TNSPNT(3) = TNSPNT3
    TNSPNT(4) = TNSPNT4
    TNSPNT(5) = TNSPNT5
    TNSPNT(6) = TNSPNT6
    TNSPNT(7) = TNSPNT7
    TNSPNT(8) = TNSPNT8
    TNSPNT(9) = TNSPNT9
    
    For i = 0 To 9
        OutDspWave(i).CreateConstant 0, 6
        vIndex = TNSPNT(i)
        For j = 0 To 5
            If vIndex > 0 Then
                OutDspWave(i).Element(j) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDspWave(i).Element(j) = 0
                
            End If
    
        Next j
    Next i
    
    OutDSPWave0 = OutDspWave(0).Copy
    OutDSPWave1 = OutDspWave(1).Copy
    OutDSPWave2 = OutDspWave(2).Copy
    OutDSPWave3 = OutDspWave(3).Copy
    OutDSPWave4 = OutDspWave(4).Copy
    OutDSPWave5 = OutDspWave(5).Copy
    OutDSPWave6 = OutDspWave(6).Copy
    OutDSPWave7 = OutDspWave(7).Copy
    OutDSPWave8 = OutDspWave(8).Copy
    OutDSPWave9 = OutDspWave(9).Copy
    

End Function


Public Function NANDIO_CAL_UD(InDSPWave As DSPWave, OutDSPWave0 As DSPWave, OutDSPWave1 As DSPWave, OutDSPWave2 As DSPWave, OutDSPWave3 As DSPWave, OutDSPWave4 As DSPWave, OutDSPWave5 As DSPWave, OutDSPWave6 As DSPWave, OutDSPWave7 As DSPWave, OutDSPWave8 As DSPWave, OutDSPWave9 As DSPWave, _
                                TNSPNT0 As Long, TNSPNT1 As Long, TNSPNT2 As Long, TNSPNT3 As Long, TNSPNT4 As Long, TNSPNT5 As Long, TNSPNT6 As Long, TNSPNT7 As Long, TNSPNT8 As Long, TNSPNT9 As Long) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
    Dim j As Integer
    Dim OutDspWave(9) As New DSPWave
    Dim TNSPNT(9) As Long
    
    For i = 0 To 9
        OutDspWave(i) = InDSPWave.Select(i, 10, 63).Copy
        index = OutDspWave(i).CalcSum
        Call OutDspWave(i).CalcMaximumValue(vIndex)
        TNSPNT(i) = vIndex
        
        If 63 - vIndex - index <> 0 Then
             For j = 0 To OutDspWave(i).SampleSize - 1
                
                If OutDspWave(i).Element(j) = 1 Then
                    index = j
                    Exit For
                End If
                
             Next j
             TNSPNT(i) = index
        End If
        
    Next i
    
    OutDSPWave0 = OutDspWave(0).Copy
    OutDSPWave1 = OutDspWave(1).Copy
    OutDSPWave2 = OutDspWave(2).Copy
    OutDSPWave3 = OutDspWave(3).Copy
    OutDSPWave4 = OutDspWave(4).Copy
    OutDSPWave5 = OutDspWave(5).Copy
    OutDSPWave6 = OutDspWave(6).Copy
    OutDSPWave7 = OutDspWave(7).Copy
    OutDSPWave8 = OutDspWave(8).Copy
    OutDSPWave9 = OutDspWave(9).Copy
    
    TNSPNT0 = TNSPNT(0)
    TNSPNT1 = TNSPNT(1)
    TNSPNT2 = TNSPNT(2)
    TNSPNT3 = TNSPNT(3)
    TNSPNT4 = TNSPNT(4)
    TNSPNT5 = TNSPNT(5)
    TNSPNT6 = TNSPNT(6)
    TNSPNT7 = TNSPNT(7)
    TNSPNT8 = TNSPNT(8)
    TNSPNT9 = TNSPNT(9)

End Function

Public Function NAND_PVTP(InDSPWave As DSPWave, PVTPNR_int As Long, OutDspWave As DSPWave, OutDSPWaveP1 As DSPWave, OutDSPWaveP2 As DSPWave, OutDSPWaveP3 As DSPWave, OutDSPWaveP4 As DSPWave, OutDSPWaveM1 As DSPWave) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
        
        OutDspWave.CreateConstant 0, 6
        OutDSPWaveP1.CreateConstant 0, 6
        OutDSPWaveP2.CreateConstant 0, 6
        OutDSPWaveP3.CreateConstant 0, 6
        OutDSPWaveP4.CreateConstant 0, 6
        OutDSPWaveM1.CreateConstant 0, 6
        
        index = InDSPWave.CalcSum
        Call InDSPWave.CalcMaximumValue(vIndex)
        PVTPNR_int = 31 - vIndex
        
        If 32 - vIndex - index <> 0 Then
             For i = 0 To InDSPWave.SampleSize - 1
                
                If InDSPWave.Element(i) = 1 Then
                    index = i
                    Exit For
                End If
                
             Next i
             PVTPNR_int = 31 - index
        End If
        
        vIndex = PVTPNR_int
        For i = 0 To 5
            If vIndex > 0 Then
                OutDspWave.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDspWave.Element(i) = 0
                
            End If

        Next i
        
        vIndex = PVTPNR_int + 1
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveP1.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveP1.Element(i) = 0
                
            End If

        Next i

        vIndex = PVTPNR_int + 2
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveP2.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveP2.Element(i) = 0
                
            End If

        Next i

        vIndex = PVTPNR_int + 3
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveP3.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveP3.Element(i) = 0
                
            End If

        Next i

        vIndex = PVTPNR_int + 4
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveP4.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveP4.Element(i) = 0
                
            End If

        Next i

        vIndex = PVTPNR_int - 1
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveM1.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveM1.Element(i) = 0
                
            End If

        Next i


End Function
Public Function NAND_PVTN(InDSPWave As DSPWave, PVTPNR_int As Long, OutDspWave As DSPWave, OutDSPWaveP1 As DSPWave, OutDSPWaveP2 As DSPWave, OutDSPWaveP3 As DSPWave, OutDSPWaveP4 As DSPWave, OutDSPWaveM1 As DSPWave) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
        
        OutDspWave.CreateConstant 0, 6
        OutDSPWaveP1.CreateConstant 0, 6
        OutDSPWaveP2.CreateConstant 0, 6
        OutDSPWaveP3.CreateConstant 0, 6
        OutDSPWaveP4.CreateConstant 0, 6
        OutDSPWaveM1.CreateConstant 0, 6
        
        index = InDSPWave.CalcSum
        Call InDSPWave.CalcMaximumValue(vIndex)
        PVTPNR_int = vIndex
        
        If 32 - PVTPNR_int - index <> 0 Then
             For i = 0 To InDSPWave.SampleSize - 1
                
                If InDSPWave.Element(i) = 1 Then
                    index = i
                    Exit For
                End If
                
             Next i
             PVTPNR_int = index
        End If
        
        vIndex = PVTPNR_int
        For i = 0 To 5
            If vIndex > 0 Then
                OutDspWave.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
            Else
                OutDspWave.Element(i) = 0
            End If
        
        Next i
        
        vIndex = PVTPNR_int + 1
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveP1.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveP1.Element(i) = 0
                
            End If

        Next i

        vIndex = PVTPNR_int + 2
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveP2.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveP2.Element(i) = 0
                
            End If

        Next i

        vIndex = PVTPNR_int + 3
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveP3.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveP3.Element(i) = 0
                
            End If

        Next i

        vIndex = PVTPNR_int + 4
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveP4.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveP4.Element(i) = 0
                
            End If

        Next i

        vIndex = PVTPNR_int - 1
        For i = 0 To 5
            If vIndex > 0 Then
                OutDSPWaveM1.Element(i) = vIndex Mod 2
                vIndex = Int(vIndex / 2)
                
            Else
                OutDSPWaveM1.Element(i) = 0
                
            End If

        Next i
                
        
End Function
Public Function NAND_PVTR(InDSPWave As DSPWave, PVTPNR_int As Long, OutDspWave As DSPWave, OutDSPWave2 As DSPWave) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
    Dim x As Integer
        
        OutDspWave.CreateConstant 0, 15
        OutDSPWave2.CreateConstant 0, 5
        
        index = 32 - InDSPWave.CalcSum
        Call InDSPWave.CalcMinimumValue(vIndex)
        PVTPNR_int = vIndex
        
        If 32 - PVTPNR_int - index <> 0 Then
             For i = 0 To InDSPWave.SampleSize - 1
                
                If InDSPWave.Element(i) = 0 Then
                    index = i
                    Exit For
                End If
                PVTPNR_int = index
                
             Next i
        
        End If
        
        vIndex = PVTPNR_int
'        For i = 14 To 0 Step -3
'            If vIndex > 0 Then
'                x = vIndex Mod 2
'                OutDSPWave2.Element(((i + 1) / 3) - 1) = x
'                OutDspWave.Element(i) = x
'                OutDspWave.Element(i - 1) = x
'                OutDspWave.Element(i - 2) = x
'                vIndex = Int(vIndex / 2)
'            Else
'                OutDSPWave2.Element(((i + 1) / 3) - 1) = 0
'                OutDspWave.Element(i) = 0
'                OutDspWave.Element(i - 1) = 0
'                OutDspWave.Element(i - 2) = 0
'            End If
'        Next i
'

        For i = 0 To 14 Step 3  '20170920 CT not to reverse PVTR wave
            If vIndex > 0 Then
                x = vIndex Mod 2
                OutDSPWave2.Element(i / 3) = x
                OutDspWave.Element(i) = x
                OutDspWave.Element(i + 1) = x
                OutDspWave.Element(i + 2) = x
                vIndex = Int(vIndex / 2)
            Else
                OutDSPWave2.Element(i / 3) = 0
                OutDspWave.Element(i) = 0
                OutDspWave.Element(i + 1) = 0
                OutDspWave.Element(i + 2) = 0
            End If
        Next i
        
        
End Function

Public Function RepeatWave(InDSPWave As DSPWave, OutDspWave As DSPWave, ByRef RepNum As Long, ByRef SP As Long, ByRef EP As Long) As Long

    Dim TempWave As New DSPWave
    
    TempWave = InDSPWave.Select(SP, 1, EP - SP + 1).Copy
    OutDspWave = TempWave.Repeat(RepNum)

End Function

Public Function MergeBit(PVT_Wave As DSPWave, PVTP_Wave As DSPWave, PVTN_Wave As DSPWave, PVTR_Wave As DSPWave, PVTP_P1 As DSPWave, PVTP_P2 As DSPWave, PVTP_P3 As DSPWave, PVTP_P4 As DSPWave, PVTP_M1 As DSPWave, _
PVTN_P1 As DSPWave, PVTN_P2 As DSPWave, PVTN_P3 As DSPWave, PVTN_P4 As DSPWave, PVTN_M1 As DSPWave, _
PVTPN_P1 As DSPWave, PVTPN_P2 As DSPWave, PVTPN_P3 As DSPWave, PVTPN_P4 As DSPWave, PVTP_P1_PVTN_M1 As DSPWave) As Long

    Dim TempWave(15) As New DSPWave
    Dim i As Integer
    Dim j As Integer
    For i = 0 To 15
        TempWave(i).CreateConstant 0, 17
    Next i
            
            
    For i = 0 To 16
            If i < 5 Then ''PVTR[4:0]
                For j = 0 To 15
                    'If i = 0 Then TempWave(j).CreateConstant 0, 17
                    TempWave(j).Element(i) = PVTR_Wave.Element(4 - i)  'PVT(site).Element(i) = Src2_DSPWave(site).Element(i * 3)
                Next j
            ElseIf i < 11 Then ''PVTN[5:0]
                TempWave(0).Element(i) = PVTN_Wave.Element(i - 5)      'PVT(site).Element(i) = Src1_DSPWave(site).Element(i - 5)
                TempWave(1).Element(i) = PVTN_Wave.Element(i - 5)       'PVTP_MINUS_1(site).Element(i) = Src1_DSPWave(site).Element(i - 5)
                TempWave(2).Element(i) = PVTN_M1.Element(i - 5)       'PVTN_MINUS_1(site).Element(i) = Src1_DSPWave_MINUS_1(site).Element(i - 5)
                TempWave(3).Element(i) = PVTN_Wave.Element(i - 5)       'PVTP_PLUS_1(site).Element(i) = Src1_DSPWave(site).Element(i - 5) '' 20160616 add
                TempWave(4).Element(i) = PVTN_Wave.Element(i - 5)       'PVTP_PLUS_2(site).Element(i) = Src1_DSPWave(site).Element(i - 5) '' 20160712 add
                TempWave(5).Element(i) = PVTN_Wave.Element(i - 5)       'PVTP_PLUS_3(site).Element(i) = Src1_DSPWave(site).Element(i - 5) '' 20160712 add
                TempWave(6).Element(i) = PVTN_Wave.Element(i - 5)       'PVTP_PLUS_4(site).Element(i) = Src1_DSPWave(site).Element(i - 5) '' 20160712 add
                TempWave(7).Element(i) = PVTN_P1.Element(i - 5)       'PVTN_PLUS_1(site).Element(i) = Src1_DSPWave_PLUS_1(site).Element(i - 5) '' 20160712 add
                TempWave(8).Element(i) = PVTN_P2.Element(i - 5)       'PVTN_PLUS_2(site).Element(i) = Src1_DSPWave_PLUS_2(site).Element(i - 5) '' 20160712 add
                TempWave(9).Element(i) = PVTN_P3.Element(i - 5)       'PVTN_PLUS_3(site).Element(i) = Src1_DSPWave_PLUS_3(site).Element(i - 5) '' 20160712 add
                TempWave(10).Element(i) = PVTN_P4.Element(i - 5)       'PVTN_PLUS_4(site).Element(i) = Src1_DSPWave_PLUS_4(site).Element(i - 5) '' 20160712 add
                TempWave(11).Element(i) = PVTN_P1.Element(i - 5)       'PVTPN_PLUS_1(site).Element(i) = Src1_DSPWave_PLUS_1(site).Element(i - 5) '' 20160712 add
                TempWave(12).Element(i) = PVTN_P2.Element(i - 5)       'PVTPN_PLUS_2(site).Element(i) = Src1_DSPWave_PLUS_2(site).Element(i - 5) '' 20160712 add
                TempWave(13).Element(i) = PVTN_P3.Element(i - 5)       'PVTPN_PLUS_3(site).Element(i) = Src1_DSPWave_PLUS_3(site).Element(i - 5) '' 20160712 add
                TempWave(14).Element(i) = PVTN_P4.Element(i - 5)       'PVTPN_PLUS_4(site).Element(i) = Src1_DSPWave_PLUS_4(site).Element(i - 5) '' 20160712 add
                TempWave(15).Element(i) = PVTN_M1.Element(i - 5)       'PVTP_PLUS_1_PVTN_MINUS_1(Site).Element(i) = Src1_DSPWave_MINUS_1(Site).Element(i - 5) '' 20161115 add
            Else               ''PVTP[5:0]
                TempWave(0).Element(i) = PVTP_Wave.Element(i - 11)     'PVT(site).Element(i) = Src_DSPWave(site).Element(i - 11)
                TempWave(1).Element(i) = PVTP_M1.Element(i - 11)     'PVTP_MINUS_1(site).Element(i) = Src_DSPWave_MINUS_1(site).Element(i - 11)
                TempWave(2).Element(i) = PVTP_Wave.Element(i - 11)     'PVTN_MINUS_1(site).Element(i) = Src_DSPWave(site).Element(i - 11)
                TempWave(3).Element(i) = PVTP_P1.Element(i - 11)     'PVTP_PLUS_1(site).Element(i) = Src_DSPWave_PLUS_1(site).Element(i - 11) '' 20160616 add
                TempWave(4).Element(i) = PVTP_P2.Element(i - 11)     'PVTP_PLUS_2(site).Element(i) = Src_DSPWave_PLUS_2(site).Element(i - 11) '' 20160712 add
                TempWave(5).Element(i) = PVTP_P3.Element(i - 11)     'PVTP_PLUS_3(site).Element(i) = Src_DSPWave_PLUS_3(site).Element(i - 11) '' 20160712 add
                TempWave(6).Element(i) = PVTP_P4.Element(i - 11)     'PVTP_PLUS_4(site).Element(i) = Src_DSPWave_PLUS_4(site).Element(i - 11) '' 20160712 add
                TempWave(7).Element(i) = PVTP_Wave.Element(i - 11)     'PVTN_PLUS_1(site).Element(i) = Src_DSPWave(site).Element(i - 11) '' 20160712 add
                TempWave(8).Element(i) = PVTP_Wave.Element(i - 11)     'PVTN_PLUS_2(site).Element(i) = Src_DSPWave(site).Element(i - 11) '' 20160712 add
                TempWave(9).Element(i) = PVTP_Wave.Element(i - 11)     'PVTN_PLUS_3(site).Element(i) = Src_DSPWave(site).Element(i - 11) '' 20160712 add
                TempWave(10).Element(i) = PVTP_Wave.Element(i - 11)     'PVTN_PLUS_4(site).Element(i) = Src_DSPWave(site).Element(i - 11) '' 20160712 add
                TempWave(11).Element(i) = PVTP_P1.Element(i - 11)     'PVTPN_PLUS_1(site).Element(i) = Src_DSPWave_PLUS_1(site).Element(i - 11) '' 20160712 add
                TempWave(12).Element(i) = PVTP_P2.Element(i - 11)     'PVTPN_PLUS_2(site).Element(i) = Src_DSPWave_PLUS_2(site).Element(i - 11) '' 20160712 add
                TempWave(13).Element(i) = PVTP_P3.Element(i - 11)     'PVTPN_PLUS_3(site).Element(i) = Src_DSPWave_PLUS_3(site).Element(i - 11) '' 20160712 add
                TempWave(14).Element(i) = PVTP_P4.Element(i - 11)     'PVTPN_PLUS_4(site).Element(i) = Src_DSPWave_PLUS_4(site).Element(i - 11) '' 20160712 add
                TempWave(15).Element(i) = PVTP_P1.Element(i - 11)     'PVTP_PLUS_1_PVTN_MINUS_1(Site).Element(i) = Src_DSPWave_PLUS_1(Site).Element(i - 11) '' 20160712 add
            End If
    Next i

    PVT_Wave = TempWave(0).Copy
    PVTP_M1 = TempWave(1).Copy
    PVTN_M1 = TempWave(2).Copy
    PVTP_P1 = TempWave(3).Copy
    PVTP_P2 = TempWave(4).Copy
    PVTP_P3 = TempWave(5).Copy
    PVTP_P4 = TempWave(6).Copy
    PVTN_P1 = TempWave(7).Copy
    PVTN_P2 = TempWave(8).Copy
    PVTN_P3 = TempWave(9).Copy
    PVTN_P4 = TempWave(10).Copy
    PVTPN_P1 = TempWave(11).Copy
    PVTPN_P2 = TempWave(12).Copy
    PVTPN_P3 = TempWave(13).Copy
    PVTPN_P4 = TempWave(14).Copy
    PVTP_P1_PVTN_M1 = TempWave(15).Copy
        
End Function

Public Function CAL_WAVE_CREATE(CAL_WAVE As DSPWave, CAL_DOUBLE_WAVE As DSPWave, CAL_WAVE_PVT As DSPWave, _
                                PVT_Wave As DSPWave, _
                                CAL_IO0_Wave As DSPWave, CAL_IO1_Wave As DSPWave, CAL_IO2_Wave As DSPWave, CAL_IO3_Wave As DSPWave, CAL_IO4_Wave As DSPWave, CAL_IO5_Wave As DSPWave, CAL_IO6_Wave As DSPWave, CAL_IO7_Wave As DSPWave, CAL_DQS_Wave As DSPWave, CAL_NRE_Wave As DSPWave) As Long
        
    Dim TempWave1 As New DSPWave
    Dim TempWave2 As New DSPWave
    Dim TempWave3 As New DSPWave
    Dim i As Integer
     
    '============HEADER================
     'Call MergeWave(CAL_IO0_Wave, temp(0), TempWave2) 'PVT_Wave
     'Call MergeWave(TempWave2, temp(1), TempWave2) 'CAL_NRE_Wave
   
   
   
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave2)
    Call MergeWave(TempWave2, CAL_NRE_Wave, TempWave2)
    Call MergeWave(TempWave2, PVT_Wave, TempWave2)
    Call MergeWave(TempWave2, CAL_DQS_Wave, TempWave2)
    Call MergeWave(TempWave2, PVT_Wave, TempWave2)
     
     
    '============BODY===================
    Call MergeWave(CAL_IO1_Wave, CAL_IO2_Wave, TempWave1)
    Call MergeWave(TempWave1, CAL_IO3_Wave, TempWave1)
    Call MergeWave(TempWave1, CAL_IO4_Wave, TempWave1)
    Call MergeWave(TempWave1, CAL_IO5_Wave, TempWave1)
    Call MergeWave(TempWave1, CAL_IO6_Wave, TempWave1)
    Call MergeWave(TempWave1, CAL_IO7_Wave, TempWave1)
     

    '=============CALL_WAVE==============
    Call MergeWave(TempWave2, TempWave1, CAL_WAVE)

    '=============CALL_WAVE_PVT============
    Call MergeWave(CAL_WAVE, TempWave2, CAL_WAVE_PVT)

    '============CAL_DOUBLE_WAVE===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO0_Wave, TempWave3)
    Call MergeWave(TempWave3, PVT_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_NRE_Wave, TempWave3)
    Call MergeWave(TempWave3, PVT_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_NRE_Wave, TempWave3)
    Call MergeWave(TempWave3, PVT_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_DQS_Wave, TempWave3)
    Call MergeWave(TempWave3, PVT_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_DQS_Wave, TempWave3)
    Call MergeWave(TempWave3, PVT_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO1_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO1_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO2_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO2_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO3_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO3_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO4_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO4_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO5_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO5_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO6_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO6_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO7_Wave, TempWave3)
    Call MergeWave(TempWave3, CAL_IO7_Wave, TempWave3)
    'Str_CAL_WAVE = Str_RX_CAL_IO7 & Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO5 & _
                Str_RX_CAL_IO4 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_RX_CAL_IO1 & Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_NRE & _
                Str_PVT & Str_RX_CAL_IO0 & Str_PVT & Str_RX_CAL_IO0
    CAL_DOUBLE_WAVE = TempWave3.Copy
    
End Function

Public Function CAL_WAVE_CREATE_1(CAL_WAVE_PVTP_MINUS_1 As DSPWave, CAL_WAVE_PVTP_PLUS_1 As DSPWave, CAL_WAVE_PVTP_PLUS_2 As DSPWave, CAL_WAVE_PVTP_PLUS_3 As DSPWave, CAL_WAVE_PVTP_PLUS_4 As DSPWave, _
                                PVT_Wave As DSPWave, _
                                PVTP_PLUS_1_Wave As DSPWave, PVTP_PLUS_2_Wave As DSPWave, PVTP_PLUS_3_Wave As DSPWave, PVTP_PLUS_4_Wave As DSPWave, PVTP_MINUS_1_Wave As DSPWave, _
                                CAL_IO0_Wave As DSPWave, CAL_IO1_Wave As DSPWave, CAL_IO2_Wave As DSPWave, CAL_IO3_Wave As DSPWave, CAL_IO4_Wave As DSPWave, CAL_IO5_Wave As DSPWave, CAL_IO6_Wave As DSPWave, CAL_IO7_Wave As DSPWave, CAL_DQS_Wave As DSPWave, CAL_NRE_Wave As DSPWave) As Long

    Dim TempWave4 As New DSPWave
    Dim TempWave5 As New DSPWave
    Dim TempWave6 As New DSPWave
    Dim TempWave7 As New DSPWave
    Dim TempWave8 As New DSPWave

'============CAL_WAVE_PVTP_MINUS_1===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_NRE_Wave, TempWave4)
    Call MergeWave(TempWave4, PVT_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_DQS_Wave, TempWave4)
    Call MergeWave(TempWave4, PVT_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_IO1_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_IO2_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_IO3_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_IO4_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_IO5_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_IO6_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_IO7_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_IO0_Wave, TempWave4)
    Call MergeWave(TempWave4, PVTP_MINUS_1_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_NRE_Wave, TempWave4)
    Call MergeWave(TempWave4, PVTP_MINUS_1_Wave, TempWave4)
    Call MergeWave(TempWave4, CAL_DQS_Wave, TempWave4)
    Call MergeWave(TempWave4, PVTP_MINUS_1_Wave, TempWave4)
    'Str_CAL_WAVE = Str_PVTP_MINUS_1 & Str_RX_CAL_DQS & Str_PVTP_MINUS_1 & Str_RX_CAL_NRE & Str_PVTP_MINUS_1 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTP_MINUS_1 = TempWave4.Copy
    
    
    '============CAL_WAVE_PVTP_PLUS_1===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_NRE_Wave, TempWave5)
    Call MergeWave(TempWave5, PVT_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_DQS_Wave, TempWave5)
    Call MergeWave(TempWave5, PVT_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_IO1_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_IO2_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_IO3_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_IO4_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_IO5_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_IO6_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_IO7_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_IO0_Wave, TempWave5)
    Call MergeWave(TempWave5, PVTP_PLUS_1_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_NRE_Wave, TempWave5)
    Call MergeWave(TempWave5, PVTP_PLUS_1_Wave, TempWave5)
    Call MergeWave(TempWave5, CAL_DQS_Wave, TempWave5)
    Call MergeWave(TempWave5, PVTP_PLUS_1_Wave, TempWave5)
    'Str_CAL_WAVE = Str_PVTP_PLUS_1 & Str_RX_CAL_DQS & Str_PVTP_PLUS_1 & Str_RX_CAL_NRE & Str_PVTP_PLUS_1 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTP_PLUS_1 = TempWave5.Copy
    
    
    '============CAL_WAVE_PVTP_PLUS_2===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_NRE_Wave, TempWave6)
    Call MergeWave(TempWave6, PVT_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_DQS_Wave, TempWave6)
    Call MergeWave(TempWave6, PVT_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_IO1_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_IO2_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_IO3_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_IO4_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_IO5_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_IO6_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_IO7_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_IO0_Wave, TempWave6)
    Call MergeWave(TempWave6, PVTP_PLUS_2_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_NRE_Wave, TempWave6)
    Call MergeWave(TempWave6, PVTP_PLUS_2_Wave, TempWave6)
    Call MergeWave(TempWave6, CAL_DQS_Wave, TempWave6)
    Call MergeWave(TempWave6, PVTP_PLUS_2_Wave, TempWave6)
    'Str_CAL_WAVE = Str_PVTP_PLUS_2 & Str_RX_CAL_DQS & Str_PVTP_PLUS_2 & Str_RX_CAL_NRE & Str_PVTP_PLUS_2 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTP_PLUS_2 = TempWave6.Copy
                
                
    '============CAL_WAVE_PVTP_PLUS_3===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_NRE_Wave, TempWave7)
    Call MergeWave(TempWave7, PVT_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_DQS_Wave, TempWave7)
    Call MergeWave(TempWave7, PVT_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_IO1_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_IO2_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_IO3_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_IO4_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_IO5_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_IO6_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_IO7_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_IO0_Wave, TempWave7)
    Call MergeWave(TempWave7, PVTP_PLUS_3_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_NRE_Wave, TempWave7)
    Call MergeWave(TempWave7, PVTP_PLUS_3_Wave, TempWave7)
    Call MergeWave(TempWave7, CAL_DQS_Wave, TempWave7)
    Call MergeWave(TempWave7, PVTP_PLUS_3_Wave, TempWave7)
    'Str_CAL_WAVE = Str_PVTP_PLUS_3 & Str_RX_CAL_DQS & Str_PVTP_PLUS_3 & Str_RX_CAL_NRE & Str_PVTP_PLUS_3 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTP_PLUS_3 = TempWave7.Copy
    
    
    '============CAL_WAVE_PVTP_PLUS_4===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_NRE_Wave, TempWave8)
    Call MergeWave(TempWave8, PVT_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_DQS_Wave, TempWave8)
    Call MergeWave(TempWave8, PVT_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_IO1_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_IO2_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_IO3_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_IO4_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_IO5_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_IO6_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_IO7_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_IO0_Wave, TempWave8)
    Call MergeWave(TempWave8, PVTP_PLUS_4_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_NRE_Wave, TempWave8)
    Call MergeWave(TempWave8, PVTP_PLUS_4_Wave, TempWave8)
    Call MergeWave(TempWave8, CAL_DQS_Wave, TempWave8)
    Call MergeWave(TempWave8, PVTP_PLUS_4_Wave, TempWave8)
    'Str_CAL_WAVE = Str_PVTP_PLUS_4 & Str_RX_CAL_DQS & Str_PVTP_PLUS_4 & Str_RX_CAL_NRE & Str_PVTP_PLUS_4 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTP_PLUS_4 = TempWave8.Copy
    
End Function

Public Function CAL_WAVE_CREATE_2(CAL_WAVE_PVTN_MINUS_1 As DSPWave, CAL_WAVE_PVTN_PLUS_1 As DSPWave, CAL_WAVE_PVTN_PLUS_2 As DSPWave, CAL_WAVE_PVTN_PLUS_3 As DSPWave, CAL_WAVE_PVTN_PLUS_4 As DSPWave, _
                                PVT_Wave As DSPWave, _
                                PVTN_PLUS_1_Wave As DSPWave, PVTN_PLUS_2_Wave As DSPWave, PVTN_PLUS_3_Wave As DSPWave, PVTN_PLUS_4_Wave As DSPWave, PVTN_MINUS_1_Wave As DSPWave, _
                                CAL_IO0_Wave As DSPWave, CAL_IO1_Wave As DSPWave, CAL_IO2_Wave As DSPWave, CAL_IO3_Wave As DSPWave, CAL_IO4_Wave As DSPWave, CAL_IO5_Wave As DSPWave, CAL_IO6_Wave As DSPWave, CAL_IO7_Wave As DSPWave, CAL_DQS_Wave As DSPWave, CAL_NRE_Wave As DSPWave) As Long
                                
    Dim TempWave9 As New DSPWave
    Dim TempWave10 As New DSPWave
    Dim TempWave11 As New DSPWave
    Dim TempWave12 As New DSPWave
    Dim TempWave13 As New DSPWave
                                
 '============CAL_WAVE_PVTN_MINUS_1===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_NRE_Wave, TempWave9)
    Call MergeWave(TempWave9, PVT_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_DQS_Wave, TempWave9)
    Call MergeWave(TempWave9, PVT_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_IO1_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_IO2_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_IO3_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_IO4_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_IO5_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_IO6_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_IO7_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_IO0_Wave, TempWave9)
    Call MergeWave(TempWave9, PVTN_MINUS_1_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_NRE_Wave, TempWave9)
    Call MergeWave(TempWave9, PVTN_MINUS_1_Wave, TempWave9)
    Call MergeWave(TempWave9, CAL_DQS_Wave, TempWave9)
    Call MergeWave(TempWave9, PVTN_MINUS_1_Wave, TempWave9)
    'Str_CAL_WAVE = Str_PVTN_MINUS_1 & Str_RX_CAL_DQS & Str_PVTN_MINUS_1 & Str_RX_CAL_NRE & Str_PVTN_MINUS_1 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTN_MINUS_1 = TempWave9.Copy
    
    
    '============CAL_WAVE_PVTN_PLUS_1===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_NRE_Wave, TempWave10)
    Call MergeWave(TempWave10, PVT_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_DQS_Wave, TempWave10)
    Call MergeWave(TempWave10, PVT_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_IO1_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_IO2_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_IO3_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_IO4_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_IO5_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_IO6_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_IO7_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_IO0_Wave, TempWave10)
    Call MergeWave(TempWave10, PVTN_PLUS_1_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_NRE_Wave, TempWave10)
    Call MergeWave(TempWave10, PVTN_PLUS_1_Wave, TempWave10)
    Call MergeWave(TempWave10, CAL_DQS_Wave, TempWave10)
    Call MergeWave(TempWave10, PVTN_PLUS_1_Wave, TempWave10)
    'Str_CAL_WAVE = Str_PVTN_PLUS_1 & Str_RX_CAL_DQS & Str_PVTN_PLUS_1 & Str_RX_CAL_NRE & Str_PVTN_PLUS_1 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTN_PLUS_1 = TempWave10.Copy
    
    
    '============CAL_WAVE_PVTN_PLUS_2===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_NRE_Wave, TempWave11)
    Call MergeWave(TempWave11, PVT_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_DQS_Wave, TempWave11)
    Call MergeWave(TempWave11, PVT_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_IO1_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_IO2_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_IO3_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_IO4_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_IO5_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_IO6_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_IO7_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_IO0_Wave, TempWave11)
    Call MergeWave(TempWave11, PVTN_PLUS_2_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_NRE_Wave, TempWave11)
    Call MergeWave(TempWave11, PVTN_PLUS_2_Wave, TempWave11)
    Call MergeWave(TempWave11, CAL_DQS_Wave, TempWave11)
    Call MergeWave(TempWave11, PVTN_PLUS_2_Wave, TempWave11)
    'Str_CAL_WAVE = Str_PVTN_PLUS_2 & Str_RX_CAL_DQS & Str_PVTN_PLUS_2 & Str_RX_CAL_NRE & Str_PVTN_PLUS_2 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTN_PLUS_2 = TempWave11.Copy
    
    
    '============CAL_WAVE_PVTN_PLUS_3===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_NRE_Wave, TempWave12)
    Call MergeWave(TempWave12, PVT_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_DQS_Wave, TempWave12)
    Call MergeWave(TempWave12, PVT_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_IO1_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_IO2_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_IO3_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_IO4_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_IO5_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_IO6_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_IO7_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_IO0_Wave, TempWave12)
    Call MergeWave(TempWave12, PVTN_PLUS_3_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_NRE_Wave, TempWave12)
    Call MergeWave(TempWave12, PVTN_PLUS_3_Wave, TempWave12)
    Call MergeWave(TempWave12, CAL_DQS_Wave, TempWave12)
    Call MergeWave(TempWave12, PVTN_PLUS_3_Wave, TempWave12)
    'Str_CAL_WAVE = Str_PVTN_PLUS_3 & Str_RX_CAL_DQS & Str_PVTN_PLUS_3 & Str_RX_CAL_NRE & Str_PVTN_PLUS_3 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTN_PLUS_3 = TempWave12.Copy
    
    
    '============CAL_WAVE_PVTN_PLUS_4===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_NRE_Wave, TempWave13)
    Call MergeWave(TempWave13, PVT_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_DQS_Wave, TempWave13)
    Call MergeWave(TempWave13, PVT_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_IO1_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_IO2_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_IO3_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_IO4_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_IO5_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_IO6_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_IO7_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_IO0_Wave, TempWave13)
    Call MergeWave(TempWave13, PVTN_PLUS_4_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_NRE_Wave, TempWave13)
    Call MergeWave(TempWave13, PVTN_PLUS_4_Wave, TempWave13)
    Call MergeWave(TempWave13, CAL_DQS_Wave, TempWave13)
    Call MergeWave(TempWave13, PVTN_PLUS_4_Wave, TempWave13)
    'Str_CAL_WAVE = Str_PVTN_PLUS_4 & Str_RX_CAL_DQS & Str_PVTN_PLUS_4 & Str_RX_CAL_NRE & Str_PVTN_PLUS_4 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTN_PLUS_4 = TempWave13.Copy
                            
End Function

Public Function CAL_WAVE_CREATE_3(CAL_WAVE_PVTPN_PLUS_1 As DSPWave, CAL_WAVE_PVTPN_PLUS_2 As DSPWave, CAL_WAVE_PVTPN_PLUS_3 As DSPWave, CAL_WAVE_PVTPN_PLUS_4 As DSPWave, _
                                PVT_Wave As DSPWave, _
                                PVTPN_PLUS_1_Wave As DSPWave, PVTPN_PLUS_2_Wave As DSPWave, PVTPN_PLUS_3_Wave As DSPWave, PVTPN_PLUS_4_Wave As DSPWave, PVTP_PLUS_1_PVTN_MINUS_1_Wave As DSPWave, _
                                CAL_IO0_Wave As DSPWave, CAL_IO1_Wave As DSPWave, CAL_IO2_Wave As DSPWave, CAL_IO3_Wave As DSPWave, CAL_IO4_Wave As DSPWave, CAL_IO5_Wave As DSPWave, CAL_IO6_Wave As DSPWave, CAL_IO7_Wave As DSPWave, CAL_DQS_Wave As DSPWave, CAL_NRE_Wave As DSPWave) As Long
                                
    Dim TempWave14 As New DSPWave
    Dim TempWave15 As New DSPWave
    Dim TempWave16 As New DSPWave
    Dim TempWave17 As New DSPWave
    
 '============CAL_WAVE_PVTPN_PLUS_1===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_NRE_Wave, TempWave14)
    Call MergeWave(TempWave14, PVT_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_DQS_Wave, TempWave14)
    Call MergeWave(TempWave14, PVT_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_IO1_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_IO2_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_IO3_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_IO4_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_IO5_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_IO6_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_IO7_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_IO0_Wave, TempWave14)
    Call MergeWave(TempWave14, PVTPN_PLUS_1_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_NRE_Wave, TempWave14)
    Call MergeWave(TempWave14, PVTPN_PLUS_1_Wave, TempWave14)
    Call MergeWave(TempWave14, CAL_DQS_Wave, TempWave14)
    Call MergeWave(TempWave14, PVTPN_PLUS_1_Wave, TempWave14)
    'Str_CAL_WAVE = Str_PVTPN_PLUS_1 & Str_RX_CAL_DQS & Str_PVTPN_PLUS_1 & Str_RX_CAL_NRE & Str_PVTPN_PLUS_1 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTPN_PLUS_1 = TempWave14.Copy
    
    
    '============CAL_WAVE_PVTPN_PLUS_2===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_NRE_Wave, TempWave15)
    Call MergeWave(TempWave15, PVT_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_DQS_Wave, TempWave15)
    Call MergeWave(TempWave15, PVT_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_IO1_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_IO2_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_IO3_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_IO4_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_IO5_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_IO6_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_IO7_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_IO0_Wave, TempWave15)
    Call MergeWave(TempWave15, PVTPN_PLUS_2_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_NRE_Wave, TempWave15)
    Call MergeWave(TempWave15, PVTPN_PLUS_2_Wave, TempWave15)
    Call MergeWave(TempWave15, CAL_DQS_Wave, TempWave15)
    Call MergeWave(TempWave15, PVTPN_PLUS_2_Wave, TempWave15)
    'Str_CAL_WAVE = Str_PVTPN_PLUS_2 & Str_RX_CAL_DQS & Str_PVTPN_PLUS_2 & Str_RX_CAL_NRE & Str_PVTPN_PLUS_2 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTPN_PLUS_2 = TempWave15.Copy
    
    
    '============CAL_WAVE_PVTPN_PLUS_3===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_NRE_Wave, TempWave16)
    Call MergeWave(TempWave16, PVT_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_DQS_Wave, TempWave16)
    Call MergeWave(TempWave16, PVT_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_IO1_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_IO2_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_IO3_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_IO4_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_IO5_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_IO6_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_IO7_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_IO0_Wave, TempWave16)
    Call MergeWave(TempWave16, PVTPN_PLUS_3_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_NRE_Wave, TempWave16)
    Call MergeWave(TempWave16, PVTPN_PLUS_3_Wave, TempWave16)
    Call MergeWave(TempWave16, CAL_DQS_Wave, TempWave16)
    Call MergeWave(TempWave16, PVTPN_PLUS_3_Wave, TempWave16)
    'Str_CAL_WAVE = Str_PVTPN_PLUS_3 & Str_RX_CAL_DQS & Str_PVTPN_PLUS_3 & Str_RX_CAL_NRE & Str_PVTPN_PLUS_3 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTPN_PLUS_3 = TempWave16.Copy
    
    
    '============CAL_WAVE_PVTPN_PLUS_4===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_NRE_Wave, TempWave17)
    Call MergeWave(TempWave17, PVT_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_DQS_Wave, TempWave17)
    Call MergeWave(TempWave17, PVT_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_IO1_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_IO2_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_IO3_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_IO4_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_IO5_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_IO6_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_IO7_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_IO0_Wave, TempWave17)
    Call MergeWave(TempWave17, PVTPN_PLUS_4_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_NRE_Wave, TempWave17)
    Call MergeWave(TempWave17, PVTPN_PLUS_4_Wave, TempWave17)
    Call MergeWave(TempWave17, CAL_DQS_Wave, TempWave17)
    Call MergeWave(TempWave17, PVTPN_PLUS_4_Wave, TempWave17)
    'Str_CAL_WAVE = Str_PVTPN_PLUS_4 & Str_RX_CAL_DQS & Str_PVTPN_PLUS_4 & Str_RX_CAL_NRE & Str_PVTPN_PLUS_4 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTPN_PLUS_4 = TempWave17.Copy
                                
End Function

Public Function CAL_WAVE_CREATE_4(CAL_WAVE_PVTP_PLUS_1_PVTN_MINUS_1 As DSPWave, CAL_WAVE_PVTP_PLUS_1_PVT As DSPWave, CAL_WAVE_PVTP_PLUS_1_PVTN_MINUS_1_PVTP_PLUS_1 As DSPWave, _
                                PVT_Wave As DSPWave, _
                                PVTP_PLUS_1_Wave As DSPWave, PVTP_PLUS_1_PVTN_MINUS_1_Wave As DSPWave, _
                                CAL_IO0_Wave As DSPWave, CAL_IO1_Wave As DSPWave, CAL_IO2_Wave As DSPWave, CAL_IO3_Wave As DSPWave, CAL_IO4_Wave As DSPWave, CAL_IO5_Wave As DSPWave, CAL_IO6_Wave As DSPWave, CAL_IO7_Wave As DSPWave, CAL_DQS_Wave As DSPWave, CAL_NRE_Wave As DSPWave) As Long
    
    Dim TempWave18 As New DSPWave
    Dim TempWave19 As New DSPWave
    Dim TempWave20 As New DSPWave
    
    '============CAL_WAVE_PVTP_PLUS_1_PVTN_MINUS_1===============
    Call MergeWave(CAL_IO0_Wave, PVT_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_NRE_Wave, TempWave18)
    Call MergeWave(TempWave18, PVT_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_DQS_Wave, TempWave18)
    Call MergeWave(TempWave18, PVT_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_IO1_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_IO2_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_IO3_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_IO4_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_IO5_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_IO6_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_IO7_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_IO0_Wave, TempWave18)
    Call MergeWave(TempWave18, PVTP_PLUS_1_PVTN_MINUS_1_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_NRE_Wave, TempWave18)
    Call MergeWave(TempWave18, PVTP_PLUS_1_PVTN_MINUS_1_Wave, TempWave18)
    Call MergeWave(TempWave18, CAL_DQS_Wave, TempWave18)
    Call MergeWave(TempWave18, PVTP_PLUS_1_PVTN_MINUS_1_Wave, TempWave18)
    'Str_CAL_WAVE = Str_PVTP_PLUS_1_PVTN_MINUS_1 & Str_RX_CAL_DQS & Str_PVTP_PLUS_1_PVTN_MINUS_1 & Str_RX_CAL_NRE & Str_PVTP_PLUS_1_PVTN_MINUS_1 & Str_RX_CAL_IO0 & _
                Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & _
                Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0
    CAL_WAVE_PVTP_PLUS_1_PVTN_MINUS_1 = TempWave18.Copy
    
    
    '============CAL_WAVE_PVTP_PLUS_1_PVT===============
    Call MergeWave(CAL_IO0_Wave, PVTP_PLUS_1_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_NRE_Wave, TempWave19)
    Call MergeWave(TempWave19, PVTP_PLUS_1_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_DQS_Wave, TempWave19)
    Call MergeWave(TempWave19, PVTP_PLUS_1_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_IO1_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_IO2_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_IO3_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_IO4_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_IO5_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_IO6_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_IO7_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_IO0_Wave, TempWave19)
    Call MergeWave(TempWave19, PVT_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_NRE_Wave, TempWave19)
    Call MergeWave(TempWave19, PVT_Wave, TempWave19)
    Call MergeWave(TempWave19, CAL_DQS_Wave, TempWave19)
    Call MergeWave(TempWave19, PVT_Wave, TempWave19)
    'Str_CAL_WAVE = Str_PVT & Str_RX_CAL_DQS & Str_PVT & Str_RX_CAL_NRE & Str_PVT & Str_RX_CAL_IO0 & Str_RX_CAL_IO7 & _
                Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & Str_PVTP_PLUS_1 & _
                Str_RX_CAL_DQS & Str_PVTP_PLUS_1 & Str_RX_CAL_NRE & Str_PVTP_PLUS_1 & Str_RX_CAL_IO0
    CAL_WAVE_PVTP_PLUS_1_PVT = TempWave19.Copy
    
    
    '============CAL_WAVE_PVTP_PLUS_1_PVTN_MINUS_1_PVTP_PLUS_1===============
    Call MergeWave(CAL_IO0_Wave, PVTP_PLUS_1_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_NRE_Wave, TempWave20)
    Call MergeWave(TempWave20, PVTP_PLUS_1_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_DQS_Wave, TempWave20)
    Call MergeWave(TempWave20, PVTP_PLUS_1_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_IO1_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_IO2_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_IO3_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_IO4_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_IO5_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_IO6_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_IO7_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_IO0_Wave, TempWave20)
    Call MergeWave(TempWave20, PVTP_PLUS_1_PVTN_MINUS_1_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_NRE_Wave, TempWave20)
    Call MergeWave(TempWave20, PVTP_PLUS_1_PVTN_MINUS_1_Wave, TempWave20)
    Call MergeWave(TempWave20, CAL_DQS_Wave, TempWave20)
    Call MergeWave(TempWave20, PVTP_PLUS_1_PVTN_MINUS_1_Wave, TempWave20)
    'Str_CAL_WAVE = Str_PVTP_PLUS_1_PVTN_MINUS_1 & Str_RX_CAL_DQS & Str_PVTP_PLUS_1_PVTN_MINUS_1 & Str_RX_CAL_NRE & _
                Str_PVTP_PLUS_1_PVTN_MINUS_1 & Str_RX_CAL_IO0 & Str_RX_CAL_IO7 & Str_RX_CAL_IO6 & Str_RX_CAL_IO5 & Str_RX_CAL_IO4 & _
                Str_RX_CAL_IO3 & Str_RX_CAL_IO2 & Str_RX_CAL_IO1 & Str_PVTP_PLUS_1 & Str_RX_CAL_DQS & Str_PVTP_PLUS_1 & _
                Str_RX_CAL_NRE & Str_PVTP_PLUS_1 & Str_RX_CAL_IO0
    CAL_WAVE_PVTP_PLUS_1_PVTN_MINUS_1_PVTP_PLUS_1 = TempWave20.Copy
    
End Function


Public Function PVTPN_Wave_Gen(RefWave As DSPWave, InWave As DSPWave, outwave As DSPWave) As Long

     RefWave.Select(183, 1, 17).Replace (InWave)
     RefWave.Select(212, 1, 17).Replace (InWave)
     RefWave.Select(241, 1, 17).Replace (InWave)
     outwave = RefWave.Copy
     
End Function

Public Function PVTPN_New_Wave_Gen(RefWave As DSPWave, inWave1 As DSPWave, inWave2 As DSPWave, outwave As DSPWave) As Long

     
     RefWave.Select(12, 1, 17).Replace (inWave1)
     RefWave.Select(41, 1, 17).Replace (inWave1)
     RefWave.Select(70, 1, 17).Replace (inWave1)
     
     RefWave.Select(183, 1, 17).Replace (inWave2)
     RefWave.Select(212, 1, 17).Replace (inWave2)
     RefWave.Select(241, 1, 17).Replace (inWave2)
     
     outwave = RefWave.Copy
     
End Function

Public Function PassWave(InWave As DSPWave, outwave As DSPWave) As Long
    outwave = InWave.Copy
End Function


Public Function Calc_Average_DSP_Porcedure(InWf1 As DSPWave, InWf2 As DSPWave, outwave As DSPWave, Average_Out As Double) As Long

    InWf1 = InWf1.ConvertDataTypeTo(DspLong)
    InWf2 = InWf2.ConvertDataTypeTo(DspLong)
    InWf1 = InWf1.ConvertStreamTo(tldspParallel, InWf1.SampleSize, 0, Bit0IsMsb)
    InWf2 = InWf2.ConvertStreamTo(tldspParallel, InWf2.SampleSize, 0, Bit0IsMsb)
    outwave = InWf1.Add(InWf2)
    Average_Out = outwave.Element(0) / 2
'    Average_Out = Average_Out / 2
    
End Function



Public Function Calc_NAND_PHY_MDLL_DSP(InWf1 As DSPWave, InWf2 As DSPWave, InWf3 As DSPWave, InWf4 As DSPWave, OutWave1 As DSPWave, Result_Ratio As Double) As Long
    Dim i As Integer
    Dim DSPWave_Combine_Dec(3) As New DSPWave
    
    For i = 0 To 3
        DSPWave_Combine_Dec(i).CreateConstant 0, 1, DspLong
    Next i
'    Call ConvertToLongAndSerialToParrel(InWf1, 9, DSPWave_Combine_Dec(0))
'    Call ConvertToLongAndSerialToParrel(InWf2, 9, DSPWave_Combine_Dec(1))
'    Call ConvertToLongAndSerialToParrel(InWf3, 9, DSPWave_Combine_Dec(2))
'    Call ConvertToLongAndSerialToParrel(InWf4, 9, DSPWave_Combine_Dec(3))
    
    OutWave1.Element(0) = DSPWave_Combine_Dec(0).Element(0) - DSPWave_Combine_Dec(1).Element(0)
    OutWave1.Element(1) = DSPWave_Combine_Dec(2).Element(0) - DSPWave_Combine_Dec(3).Element(0)
    If OutWave1.Element(0) = 0 Then OutWave1.Element(0) = 99999999
    Result_Ratio = OutWave1.Element(1) / OutWave1.Element(0)
    
End Function

Public Function Trim_ConvertToLongAndSerialToParrel(ByVal InWf As DSPWave, ByRef DataWf As DSPWave) As Long
    
    InWf = InWf.ConvertDataTypeTo(DspLong)
    If (InWf.SampleSize > 1) Then
        DataWf = InWf.ConvertStreamTo(tldspParallel, InWf.SampleSize, 0, Bit0IsMsb)
    Else
        DataWf = InWf.Copy
    End If
    
End Function

Public Function DSPWf_Merge(InA As DSPWave, InB As DSPWave, out As DSPWave) As Long

    InA = InA.ConvertDataTypeTo(DspDouble)
    InB = InB.ConvertDataTypeTo(DspDouble)
    
    out = InA.Concatenate(InB)

End Function

Public Function RPIndexSetting(InA As DSPWave, InB As DSPWave, ByVal InRP As Double) As Long
Dim A_sample As New DSPWave
Dim B_sample As Double
    A_sample.CreateConstant 0, 1
    A_sample.Element(0) = InA.SampleSize + InRP
    Call DSPWf_Merge(InB, A_sample, InB)
End Function

Public Function TrimSrcRP(InA As DSPWave, RP As DSPWave, TrimSrc As DSPWave, OutWf As DSPWave) As Long

Dim i As Long
Dim index As Double
Dim B As Double
    OutWf = InA
    For i = 1 To RP.SampleSize
        B = RP.Element(i - 1) - TrimSrc.SampleSize
        If B > 0.09 And B <= 0.11 Then
            OutWf.Select(index, 1, TrimSrc.SampleSize).Replace (TrimSrc)
            index = index + RP.Element(i - 1) - 0.1
        Else
            index = index + RP.Element(i - 1)
        End If
    Next i
End Function

