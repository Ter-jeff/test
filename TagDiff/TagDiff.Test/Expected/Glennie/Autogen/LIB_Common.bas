#Const isUFP = False
Attribute VB_Name = "LIB_Common"
Public Function setXY(x As Integer, y As Integer, Optional Device As String) As Long

    Dim funcName As String:: funcName = "setXY"
    
    Dim m_chmapName As String
    Dim m_siteCnt As Long
    Dim m_match_flag As Boolean
    
    m_match_flag = False
    m_siteCnt = TheExec.Sites.Existing.count
    m_chmapName = LCase(TheExec.CurrentChanMap)
    
    If (m_siteCnt = 1) Then
        m_match_flag = True
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)

    ElseIf (m_siteCnt = 2) Then
        If (m_chmapName Like "*ch*2*") Then
            m_match_flag = True
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, x + 3)
            
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y)
            
            Debug.Print "Site0: " & "(" & x & "," & y & ")"
            Debug.Print "Site1: " & "(" & x + 3 & "," & y & ")"
            
        End If
    ElseIf (m_siteCnt = 3) Then
        If (m_chmapName Like "*ch*3*") Then
            m_match_flag = True
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, x + 3)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, x + 6)
            
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y)

            Debug.Print "Site0: " & "(" & x & "," & y & ")"
            Debug.Print "Site1: " & "(" & x + 3 & "," & y & ")"
            Debug.Print "Site2: " & "(" & x + 6 & "," & y & ")"


        End If




    ElseIf (m_siteCnt = 4) Then
        If (m_chmapName Like "*ch*4*") Then
            m_match_flag = True
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, x + 3) '20221220
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, x + 3) '20221220
            
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 3) '20221220
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 3) '20221220
            
            Debug.Print "Site0: " & "(" & x & "," & y & ")"
            Debug.Print "Site1: " & "(" & x & "," & y - 3 & ")"
            Debug.Print "Site2: " & "(" & x + 3 & "," & y & ")"
            Debug.Print "Site3: " & "(" & x + 3 & "," & y - 3 & ")"
        End If

    ElseIf (m_siteCnt = 6) Then
        If (m_chmapName Like "*ch*6*") Then
            m_match_flag = True
             If UCase(m_chmapName) Like "*CP*" Then
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, x + 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, x + 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(4, x + 4)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(5, x + 4)
                
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 4)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 4)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 4)
                
                Debug.Print "Site0: " & "(" & x & "," & y & ")"
                Debug.Print "Site1: " & "(" & x & "," & y - 4 & ")"
                Debug.Print "Site2: " & "(" & x + 2&; "," & y & ")"
                Debug.Print "Site3: " & "(" & x + 2&; "," & y - 4 & ")"
                Debug.Print "Site4: " & "(" & x + 4&; "," & y & ")"
                Debug.Print "Site5: " & "(" & x + 4&; "," & y - 4 & ")"
            ElseIf (UCase(m_chmapName) Like "*WLFT*") Then
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, x + 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, x + 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(4, x + 4)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(5, x + 4)
                
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 2)
                
                Debug.Print "Site0: " & "(" & x & "," & y & ")"
                Debug.Print "Site1: " & "(" & x & "," & y - 2 & ")"
                Debug.Print "Site2: " & "(" & x + 2&; "," & y & ")"
                Debug.Print "Site3: " & "(" & x + 2&; "," & y - 2&; ")"
                Debug.Print "Site4: " & "(" & x + 4&; "," & y & ")"
                Debug.Print "Site5: " & "(" & x + 4&; "," & y - 2 & ")"
            
            
            End If
            
        End If
    
    ElseIf (m_siteCnt = 8) Then
        If (m_chmapName Like "*ch*8*") Then
            m_match_flag = True
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, x)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, x)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, x)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(4, x + 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(5, x + 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(6, x + 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(7, x + 2)
            
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y - 4)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 6)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(6, y - 4)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(7, y - 6)
        End If

    Else
        m_match_flag = False
    End If
    
    If (m_match_flag = False) Then
        ''''Has the reminder for user to maintain this fuction if the setup is unsuitable.
        If isDebugMode Then TheExec.AddOutput "<WARNING> " + funcName + ":: The Condition Setup is Wrong."
    End If

End Function
