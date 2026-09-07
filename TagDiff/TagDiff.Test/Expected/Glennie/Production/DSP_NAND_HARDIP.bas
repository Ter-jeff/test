Attribute VB_Name = "DSP_NAND_HARDIP"

Public Function MergeWave(InPutWaveA As DSPWave, InPutWaveB As DSPWave, Outputwave As DSPWave) As Long

    Dim TempWave As New DSPWave
    
    TempWave = InPutWaveA.Concatenate(InPutWaveB)
    Outputwave = TempWave.COPY

End Function
  

Public Function NANDIO_DigCap(InDSPwave As DSPWave, PVTPNR_int As Long, OutDspWave As DSPWave, WaveStr As Double) As Long  '(CUS_Str_DigCapData As String, InDspWave As DSPWave, PVTPNR_Int As Long, OutDSPWave As DSPWave) As Long
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
        index = InDSPwave.CalcSum
        Call InDSPwave.CalcMinimumValue(vIndex)
        PVTPNR_int = vIndex
        
        If 32 - PVTPNR_int - index <> 0 Then
             For i = 0 To InDSPwave.SampleSize - 1
                
                If InDSPwave.Element(i) = 1 Then
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
            
        OutDspWave = InDSPwave.Select(0, 7, 10).COPY
            
            
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
        
    OutDSPWave0 = OutDspWave(0).COPY
    OutDSPWave1 = OutDspWave(1).COPY
    OutDSPWave2 = OutDspWave(2).COPY
    OutDSPWave3 = OutDspWave(3).COPY
    OutDSPWave4 = OutDspWave(4).COPY
    OutDSPWave5 = OutDspWave(5).COPY
    OutDSPWave6 = OutDspWave(6).COPY
    OutDSPWave7 = OutDspWave(7).COPY
    OutDSPWave8 = OutDspWave(8).COPY
    OutDSPWave9 = OutDspWave(9).COPY
    
End Function




Public Function NANDIO_CAL_DOWN_Wave(InDSPwave As DSPWave, OutDSPWave0 As DSPWave, OutDSPWave1 As DSPWave, OutDSPWave2 As DSPWave, OutDSPWave3 As DSPWave, OutDSPWave4 As DSPWave, OutDSPWave5 As DSPWave, OutDSPWave6 As DSPWave, OutDSPWave7 As DSPWave, OutDSPWave8 As DSPWave, OutDSPWave9 As DSPWave, _
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
        
    OutDSPWave0 = OutDspWave(0).COPY
    OutDSPWave1 = OutDspWave(1).COPY
    OutDSPWave2 = OutDspWave(2).COPY
    OutDSPWave3 = OutDspWave(3).COPY
    OutDSPWave4 = OutDspWave(4).COPY
    OutDSPWave5 = OutDspWave(5).COPY
    OutDSPWave6 = OutDspWave(6).COPY
    OutDSPWave7 = OutDspWave(7).COPY
    OutDSPWave8 = OutDspWave(8).COPY
    OutDSPWave9 = OutDspWave(9).COPY
    
    

End Function

Public Function NANDIO_CAL_UP_Wave(InDSPwave As DSPWave, OutDSPWave0 As DSPWave, OutDSPWave1 As DSPWave, OutDSPWave2 As DSPWave, OutDSPWave3 As DSPWave, OutDSPWave4 As DSPWave, OutDSPWave5 As DSPWave, OutDSPWave6 As DSPWave, OutDSPWave7 As DSPWave, OutDSPWave8 As DSPWave, OutDSPWave9 As DSPWave, _
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
    
    OutDSPWave0 = OutDspWave(0).COPY
    OutDSPWave1 = OutDspWave(1).COPY
    OutDSPWave2 = OutDspWave(2).COPY
    OutDSPWave3 = OutDspWave(3).COPY
    OutDSPWave4 = OutDspWave(4).COPY
    OutDSPWave5 = OutDspWave(5).COPY
    OutDSPWave6 = OutDspWave(6).COPY
    OutDSPWave7 = OutDspWave(7).COPY
    OutDSPWave8 = OutDspWave(8).COPY
    OutDSPWave9 = OutDspWave(9).COPY
    

End Function


Public Function NANDIO_CAL_UD(InDSPwave As DSPWave, OutDSPWave0 As DSPWave, OutDSPWave1 As DSPWave, OutDSPWave2 As DSPWave, OutDSPWave3 As DSPWave, OutDSPWave4 As DSPWave, OutDSPWave5 As DSPWave, OutDSPWave6 As DSPWave, OutDSPWave7 As DSPWave, OutDSPWave8 As DSPWave, OutDSPWave9 As DSPWave, _
                                TNSPNT0 As Long, TNSPNT1 As Long, TNSPNT2 As Long, TNSPNT3 As Long, TNSPNT4 As Long, TNSPNT5 As Long, TNSPNT6 As Long, TNSPNT7 As Long, TNSPNT8 As Long, TNSPNT9 As Long) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
    Dim j As Integer
    Dim OutDspWave(9) As New DSPWave
    Dim TNSPNT(9) As Long
    
    For i = 0 To 9
        OutDspWave(i) = InDSPwave.Select(i, 10, 63).COPY
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
    
    OutDSPWave0 = OutDspWave(0).COPY
    OutDSPWave1 = OutDspWave(1).COPY
    OutDSPWave2 = OutDspWave(2).COPY
    OutDSPWave3 = OutDspWave(3).COPY
    OutDSPWave4 = OutDspWave(4).COPY
    OutDSPWave5 = OutDspWave(5).COPY
    OutDSPWave6 = OutDspWave(6).COPY
    OutDSPWave7 = OutDspWave(7).COPY
    OutDSPWave8 = OutDspWave(8).COPY
    OutDSPWave9 = OutDspWave(9).COPY
    
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

Public Function NAND_PVTP(InDSPwave As DSPWave, PVTPNR_int As Long, OutDspWave As DSPWave, OutDSPWaveP1 As DSPWave, OutDSPWaveP2 As DSPWave, OutDSPWaveP3 As DSPWave, OutDSPWaveP4 As DSPWave, OutDSPWaveM1 As DSPWave) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
        
        OutDspWave.CreateConstant 0, 6
        OutDSPWaveP1.CreateConstant 0, 6
        OutDSPWaveP2.CreateConstant 0, 6
        OutDSPWaveP3.CreateConstant 0, 6
        OutDSPWaveP4.CreateConstant 0, 6
        OutDSPWaveM1.CreateConstant 0, 6
        
        index = InDSPwave.CalcSum
        Call InDSPwave.CalcMaximumValue(vIndex)
        PVTPNR_int = 31 - vIndex
        
        If 32 - vIndex - index <> 0 Then
             For i = 0 To InDSPwave.SampleSize - 1
                
                If InDSPwave.Element(i) = 1 Then
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
Public Function NAND_PVTN(InDSPwave As DSPWave, PVTPNR_int As Long, OutDspWave As DSPWave, OutDSPWaveP1 As DSPWave, OutDSPWaveP2 As DSPWave, OutDSPWaveP3 As DSPWave, OutDSPWaveP4 As DSPWave, OutDSPWaveM1 As DSPWave) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
        
        OutDspWave.CreateConstant 0, 6
        OutDSPWaveP1.CreateConstant 0, 6
        OutDSPWaveP2.CreateConstant 0, 6
        OutDSPWaveP3.CreateConstant 0, 6
        OutDSPWaveP4.CreateConstant 0, 6
        OutDSPWaveM1.CreateConstant 0, 6
        
        index = InDSPwave.CalcSum
        Call InDSPwave.CalcMaximumValue(vIndex)
        PVTPNR_int = vIndex
        
        If 32 - PVTPNR_int - index <> 0 Then
             For i = 0 To InDSPwave.SampleSize - 1
                
                If InDSPwave.Element(i) = 1 Then
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
Public Function NAND_PVTR(InDSPwave As DSPWave, PVTPNR_int As Long, OutDspWave As DSPWave, OutDSPWave2 As DSPWave) As Long
    Dim index As Integer
    Dim vIndex As Long
    Dim i As Integer
    Dim x As Integer
        
        OutDspWave.CreateConstant 0, 15
        OutDSPWave2.CreateConstant 0, 5
        
        index = 32 - InDSPwave.CalcSum
        Call InDSPwave.CalcMinimumValue(vIndex)
        PVTPNR_int = vIndex
        
        If 32 - PVTPNR_int - index <> 0 Then
             For i = 0 To InDSPwave.SampleSize - 1
                
                If InDSPwave.Element(i) = 0 Then
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

Public Function RepeatWave(InDSPwave As DSPWave, OutDspWave As DSPWave, ByRef RepNum As Long, ByRef SP As Long, ByRef EP As Long) As Long

    Dim TempWave As New DSPWave
    
    TempWave = InDSPwave.Select(SP, 1, EP - SP + 1).COPY
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

    PVT_Wave = TempWave(0).COPY
    PVTP_M1 = TempWave(1).COPY
    PVTN_M1 = TempWave(2).COPY
    PVTP_P1 = TempWave(3).COPY
    PVTP_P2 = TempWave(4).COPY
    PVTP_P3 = TempWave(5).COPY
    PVTP_P4 = TempWave(6).COPY
    PVTN_P1 = TempWave(7).COPY
    PVTN_P2 = TempWave(8).COPY
    PVTN_P3 = TempWave(9).COPY
    PVTN_P4 = TempWave(10).COPY
    PVTPN_P1 = TempWave(11).COPY
    PVTPN_P2 = TempWave(12).COPY
    PVTPN_P3 = TempWave(13).COPY
    PVTPN_P4 = TempWave(14).COPY
    PVTP_P1_PVTN_M1 = TempWave(15).COPY
        
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
    CAL_DOUBLE_WAVE = TempWave3.COPY
    
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
    CAL_WAVE_PVTP_MINUS_1 = TempWave4.COPY
    
    
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
    CAL_WAVE_PVTP_PLUS_1 = TempWave5.COPY
    
    
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
    CAL_WAVE_PVTP_PLUS_2 = TempWave6.COPY
                
                
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
    CAL_WAVE_PVTP_PLUS_3 = TempWave7.COPY
    
    
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
    CAL_WAVE_PVTP_PLUS_4 = TempWave8.COPY
    
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
    CAL_WAVE_PVTN_MINUS_1 = TempWave9.COPY
    
    
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
    CAL_WAVE_PVTN_PLUS_1 = TempWave10.COPY
    
    
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
    CAL_WAVE_PVTN_PLUS_2 = TempWave11.COPY
    
    
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
    CAL_WAVE_PVTN_PLUS_3 = TempWave12.COPY
    
    
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
    CAL_WAVE_PVTN_PLUS_4 = TempWave13.COPY
                            
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
    CAL_WAVE_PVTPN_PLUS_1 = TempWave14.COPY
    
    
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
    CAL_WAVE_PVTPN_PLUS_2 = TempWave15.COPY
    
    
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
    CAL_WAVE_PVTPN_PLUS_3 = TempWave16.COPY
    
    
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
    CAL_WAVE_PVTPN_PLUS_4 = TempWave17.COPY
                                
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
    CAL_WAVE_PVTP_PLUS_1_PVTN_MINUS_1 = TempWave18.COPY
    
    
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
    CAL_WAVE_PVTP_PLUS_1_PVT = TempWave19.COPY
    
    
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
    CAL_WAVE_PVTP_PLUS_1_PVTN_MINUS_1_PVTP_PLUS_1 = TempWave20.COPY
    
End Function


Public Function PVTPN_Wave_Gen(RefWave As DSPWave, InWave As DSPWave, outwave As DSPWave) As Long

     RefWave.Select(183, 1, 17).Replace (InWave)
     RefWave.Select(212, 1, 17).Replace (InWave)
     RefWave.Select(241, 1, 17).Replace (InWave)
     outwave = RefWave.COPY
     
End Function

Public Function PVTPN_New_Wave_Gen(RefWave As DSPWave, inWave1 As DSPWave, inWave2 As DSPWave, outwave As DSPWave) As Long

     
     RefWave.Select(12, 1, 17).Replace (inWave1)
     RefWave.Select(41, 1, 17).Replace (inWave1)
     RefWave.Select(70, 1, 17).Replace (inWave1)
     
     RefWave.Select(183, 1, 17).Replace (inWave2)
     RefWave.Select(212, 1, 17).Replace (inWave2)
     RefWave.Select(241, 1, 17).Replace (inWave2)
     
     outwave = RefWave.COPY
     
End Function

Public Function PassWave(InWave As DSPWave, outwave As DSPWave) As Long
    outwave = InWave.COPY
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

