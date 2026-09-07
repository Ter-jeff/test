Attribute VB_Name = "DSP_RTOS_UART"
Option Explicit
Public Function ProcessDUTResponse(DUTResponse As DSPWave, result As Long, SPINorATE As Long, ByVal CmdLength As Long, ByVal UseCmdLength As Boolean) As Long


    Dim PassSeq As New DSPWave
    Dim FailSeq As New DSPWave
    Dim PromptSeq As New DSPWave
    
    Dim CompareArray() As Long
    Dim PassArray(3) As Long
    Dim FailArray(3) As Long
    
    Dim i As Long
    Dim Prompt_Check As Long
    Dim ATE_prompt As Long: ATE_prompt = 0
    Dim SPIN_prompt As Long: SPIN_prompt = 1
    Dim PromptIndex As Long
    Dim EndIndex As Long

    result = 3 'Prompt is initially not found when Result stays 3
    
    'For Prompt_Check = ATE_prompt To SPIN_prompt
    
    'If Prompt_Check = SPIN_prompt Then
    ReDim CompareArray(3)
    If SPINorATE = 0 Then
        CompareArray(0) = 83 'S
        CompareArray(1) = 80 'P
        CompareArray(2) = 73 'I
        CompareArray(3) = 78 'N
    ElseIf SPINorATE = 2 Then
        ReDim CompareArray(6)
        CompareArray(0) = 83
        CompareArray(1) = 117
        CompareArray(2) = 99
        CompareArray(3) = 99
        CompareArray(4) = 101
        CompareArray(5) = 115
        CompareArray(6) = 115
        
'        CompareArray(0) = 98
'        CompareArray(1) = 97
'        CompareArray(2) = 110
'        CompareArray(3) = 107
    Else
        CompareArray(0) = 65 'A
        CompareArray(1) = 84 'T
        CompareArray(2) = 69 'E
        CompareArray(3) = 62 '>
    End If
 
    PromptSeq.data = CompareArray
    UseCmdLength = False
    If UseCmdLength Then
        If DUTResponse.SampleSize < CmdLength + 60 Then
            EndIndex = 0
        Else
            EndIndex = DUTResponse.SampleSize - (CmdLength + 60)
                        'EndIndex = 0
        End If
    Else
        EndIndex = 0
    End If
    
    If SPINorATE = 2 Then
        For i = DUTResponse.SampleSize - 6 To EndIndex Step -1
            If (DUTResponse.Select(i, 1, 7).Subtract(PromptSeq).CalcSum = 0) Then
                If (DUTResponse.Select(i, 1, 7).COPY.Subtract(PromptSeq).CalcMinimumValue > -1) Then
                    result = 2 'Result turns 2 when ATE prompt is found
                    PromptIndex = i
                    i = EndIndex
                End If
            End If
        Next i
    Else
        For i = DUTResponse.SampleSize - 3 To EndIndex Step -1
            If (DUTResponse.Select(i, 1, 4).Subtract(PromptSeq).CalcSum = 0) Then
                If (DUTResponse.Select(i, 1, 4).COPY.Subtract(PromptSeq).CalcMinimumValue > -1) Then
                    result = 2 'Result turns 2 when ATE prompt is found
                    PromptIndex = i
                    i = EndIndex
                End If
            End If
        Next i
    End If
 
'Prompt is found
    'If Result = 2 Then  'And Prompt_Check = ATE_prompt Then
    If result = 2 And SPINorATE = 1 Then
        'PassArray(0) = 10 'CR TTR-255 -->Cebu V09A new binary out short UART but different char at this position for scenarios.
        PassArray(0) = 80 'P
        PassArray(1) = 65 'A
        PassArray(2) = 83 'S
        PassArray(3) = 83 'S
        'PassArray(5) = 32 'Space TTR-13 -->Cebu V09A new binary out short UART but different char at this position for scenarios.
        
        PassSeq.data = PassArray
        
    'Fail
        'PassArray(0) = 10 'CR TTR-255 -->Cebu V09A new binary out short UART but different char at this position for scenarios.
        FailArray(0) = 70 'F
        FailArray(1) = 65 'A
        FailArray(2) = 73 'I
        FailArray(3) = 76 'L
        'PassArray(5) = 32 'Space TTR-13 -->Cebu V09A new binary out short UART but different char at this position for scenarios.
       
        FailSeq.data = FailArray
    
        For i = PromptIndex To EndIndex Step -1
    '    For i = CharCount - 10 To 0 Step -1
        
    '    If i = 115 Then
    '        Result = 5
    '    End If
            If (DUTResponse.Select(i, 1, 4).Subtract(PassSeq).CalcSum = 0) Then
                If (DUTResponse.Select(i, 1, 4).COPY.Subtract(PassSeq).CalcMinimumValue > -1) Then
                
                    result = 1 ' pass to 1
                    i = EndIndex
                End If
            End If
            
            If (DUTResponse.Select(i, 1, 4).Subtract(FailSeq).CalcSum = 0) Then
                If (DUTResponse.Select(i, 1, 4).COPY.Subtract(FailSeq).CalcMinimumValue > -1) Then
                    result = 0 ' fail to 0
                    i = EndIndex
                End If
            End If
        Next i
'  ElseIf Result = 2 And Prompt_Check = SPIN_prompt Then
    ElseIf result = 2 And (SPINorATE = 0 Or SPINorATE = 2) Then
        result = 1 'Only SPIN prompt is found
    End If
'Next Prompt_Check


End Function

Public Function CheckBootStatus(DUTResponse As DSPWave, result As Long) As Long


    Dim PassSeq As New DSPWave
    Dim FailSeq As New DSPWave
    
    Dim CompareArray(3) As Long
    Dim i As Long

'Pass
'    CompareArray(0) = 83 'S
'    CompareArray(1) = 76 'L
'    CompareArray(2) = 84 'T
'    CompareArray(3) = 62 '>

    CompareArray(0) = 65 'A
    CompareArray(1) = 84 'T
    CompareArray(2) = 69 'E
    CompareArray(3) = 62 '>
    
    PassSeq.data = CompareArray
    
'Fail
    CompareArray(0) = 70 'F
    CompareArray(1) = 65 'A
    CompareArray(2) = 73 'I
    CompareArray(3) = 76 'L
    
    FailSeq.data = CompareArray
    result = 2

    For i = DUTResponse.SampleSize - 3 To 600 Step -1   'TTR-650
'    For i = CharCount - 10 To 0 Step -1
    
'    If i = 115 Then
'        Result = 5
'    End If
        If (DUTResponse.Select(i, 1, 4).Subtract(PassSeq).CalcSum = 0) And (DUTResponse.Select(i, 1, 4).COPY.Subtract(PassSeq).CalcMinimumValue > -1) Then
            result = 1
            i = 0
        End If
'        If DUTResponse.Select(i, 1, 4).Subtract(FailSeq).CalcSum = 0 Then
'            Result = 1
'            i = 0
'        End If
    Next i


End Function
