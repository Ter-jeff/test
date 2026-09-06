Attribute VB_Name = "VBT_ZeFuseOffline"
Option Explicit
''''' fake to be the dssc result in serial, no double bits.
''''' this function is use for offline, so we don't care about test time. put it here instead of DSP
Public Function FakeBankDsscResult_Serial(bank As eFuseBdfBank, v As DataRetrieve, Optional forceGet As Boolean = False, Optional earlyfuse As Boolean = False) As DSPWave
On Error GoTo errHandler
Dim myDspwave As New DSPWave, myDspwave2 As New DSPWave, i As Long, SampleSize As Long

    Set FakeBankDsscResult_Serial = bank.GetDspWave(v, forceGet:=True, earlyfuse:=earlyfuse)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOffline", "FakeBankDsscResult_Serial")
    If AbortTest Then Exit Function Else Resume Next
End Function

''''' fake to be the dssc result in either parallel or serial, depends on the bank access type automatically
''''' this function is use for offline, so we don't care about test time. put it here instead of DSP
Public Function FakeBankDsscResult(bank As eFuseBdfBank, v As DataRetrieve, Optional forceGet As Boolean = False, Optional ForceDoubleBit As Boolean = False, Optional includeEarlyFuse As Boolean = False, Optional earlyfuse As Boolean = True) As DSPWave
On Error GoTo errHandler
Dim myDspwave As New DSPWave, myDspwave2 As New DSPWave, i As Long, j As Long, SampleSize As Long, filename As String, getPreStageWave As Boolean
Dim nBits As Long
Dim idxWave As New DSPWave
Dim site As Variant

    If GlbUtility.testedStages.Count <> 0 And Not bank.AnyWriteRun Then
        If bank.CreateFakeValueTime = 0 And (Not PseudoFuseEnable) Then
            bank.OfflinePutTestedStageValues
            bank.CreateFakeValueTime = bank.CreateFakeValueTime + 1
            If UCase(bank.name) = "ECID" Then
                bank.OfflinePutTestedStageEcidFields
            End If
        End If
        
        Set myDspwave = bank.GetOfflineDspWave(v, forceGet:=True, includeEarlyFuse:=includeEarlyFuse, earlyfuse:=earlyfuse)
    Else
        Set myDspwave = bank.GetDspWave(v, forceGet:=True, includeEarlyFuse:=includeEarlyFuse, earlyfuse:=earlyfuse)
    End If

    nBits = IIf(bank.DoubleBits, 2, 1)
    If bank.pgmMode = pgm_DAA Then
        Select Case bank.CapturePinCnt
        Case 32:
            For Each site In TheExec.sites
                myDspwave = myDspwave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, DaaParaBits / nBits, 0, Bit0IsMsb)
                If bank.DoubleBits Then
                    SampleSize = myDspwave.SampleSize
                    myDspwave2.CreateConstant 0, SampleSize
                    For i = 0 To (SampleSize - 1)
                        myDspwave2.ElementLite(i) = myDspwave.ElementLite(i) * 2# ^ 16 + myDspwave.ElementLite(i)
                    Next i
                Else
                    Set myDspwave2 = myDspwave
                End If
            Next
        Case 8:
            For Each site In TheExec.sites
                myDspwave = myDspwave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, DaaParaBits, 0, Bit0IsMsb)
                If bank.DoubleBits Then
                    SampleSize = myDspwave.SampleSize * nBits
                    If myDspwave.SampleSize Mod 2 = 1 Then
                        SampleSize = SampleSize + (2 - myDspwave.SampleSize Mod 2) * nBits
                    End If
                    myDspwave2.CreateConstant 0, SampleSize
                    For i = 0 To (SampleSize) / 4 - 1
                        myDspwave2.ElementLite(4 * i) = myDspwave.ElementLite(2 * i)
                        If (2 * i + 1 < myDspwave.SampleSize) Then
                            myDspwave2.ElementLite(4 * i + 1) = myDspwave.ElementLite(2 * i + 1)
                        End If
                        myDspwave2.ElementLite(4 * i + 2) = myDspwave.ElementLite(2 * i)
                        If (2 * i + 1 < myDspwave.SampleSize) Then
                            myDspwave2.ElementLite(4 * i + 3) = myDspwave.ElementLite(2 * i + 1)
                        End If
                    Next i
                 Else
                    Set myDspwave2 = myDspwave
                 End If
             Next
        Case 16:
            For Each site In TheExec.sites
                SampleSize = myDspwave.SampleSize
                myDspwave2.CreateConstant 0, SampleSize * 2
                For i = 0 To (SampleSize - 1)
                    j = Floor(i / 16)
                    myDspwave2.ElementLite(i + j * 16) = myDspwave.ElementLite(i)
                    myDspwave2.ElementLite(i + 16 + j * 16) = myDspwave.ElementLite(i)
                Next i
                myDspwave2 = myDspwave2.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, bank.CapturePinCnt, 0, Bit0IsMsb)
            Next
        Case Else
            GlbUtility.MessageBox "No support parallel bits " & DaaParaBits
        End Select
    Else
        For Each site In TheExec.sites
            If bank.DoubleBits And ForceDoubleBit Then  'Typically the JTAG read is the result of OR, APB is designed for read double bits through JTAG
                SampleSize = myDspwave.SampleSize * 2
                myDspwave2.CreateConstant 0, SampleSize
                For i = 0 To (myDspwave.SampleSize) / 16 - 1
                    For j = 0 To 15
                        myDspwave2.ElementLite(32 * i + j) = myDspwave.ElementLite(16 * i + j Mod 16)
                        myDspwave2.ElementLite(32 * i + j + 16) = myDspwave.ElementLite(16 * i + j Mod 16)
                    Next j
                Next i
            Else
                If bank.JtagMsbFirst Then
                    idxWave.CreateRamp myDspwave.SampleSize - 1, -1, myDspwave.SampleSize, DspLong
                    myDspwave = myDspwave.Lookup(idxWave)
                End If
                Set myDspwave2 = myDspwave
            End If
        Next
    End If
    Set FakeBankDsscResult = myDspwave2

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOffline", "FakeBankDsscResult")
    If AbortTest Then Exit Function Else Resume Next
End Function
