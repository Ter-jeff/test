Attribute VB_Name = "LIB_MBIST"
Option Explicit

Private Type MbistParamResultSyntax
    Name As String
    PattName As String      ''''it's PatternSet name
    PattRawName As String   ''''it's .pat name (New)
End Type

Private Type MbistDebugResultSyntax
    Name As String
    PMode As String
    PMode_PattName As String    ''''it's PatternSet name
    PMode_PattRawName As String ''''it's .pat name (New)
    Block As String
    Block_PattName As String    ''''it's PatternSet name
    Block_PattRawName As String ''''it's .pat name (New)
End Type

Private Type MbistCategorySyntax
    Name As String
    ChangeFlag As String
    PMode_RAW As String
    PMode() As MbistParamResultSyntax
    Block_RAW As String
    Block() As MbistParamResultSyntax
    DebugMode() As MbistDebugResultSyntax
End Type

Public Type MbistGroupSyntax
    Category() As MbistCategorySyntax ''''using dynamic array
End Type

Public CpuMbist As MbistGroupSyntax ''''using dynamic array

''''-----------------------------------------------------------------
Private Type MbistPattResultSyntax
    Name As String
    PattName As String    ''''it's PatternSet name
    PattRawName As String ''''it's .pat name (New)
End Type

Public Type MbistPatternSyntax
    Category() As MbistPattResultSyntax ''''using dynamic array
End Type

Public CpuMbist_PMode As MbistPatternSyntax
Public CpuMbist_Block As MbistPatternSyntax

''''-----------------------------------------------------------------
Private Type VoltSetResultSyntax
    Name As String
    DCCate As String
    DCSele As String
    ACCate As String
    ACSele As String
End Type

Public Type MbistVoltSetSyntax
    Category() As VoltSetResultSyntax ''''using dynamic array
End Type

Public CpuMbist_VoltSet As MbistVoltSetSyntax

''''-----------------------------------------------------------------
''''20151102 add
Private Type PowerLoopVoltSyntax
    pwrpin As String
    V_Start As Double
    V_Stop As Double
    V_Step As Double
    Enable As String
End Type

Public Type MbistPowerLoopSyntax
    Category() As PowerLoopVoltSyntax ''''using dynamic array
End Type

Public CpuMbist_Power As MbistPowerLoopSyntax

''''-----------------------------------------------------------------
Public gB_findCpuMbist_flag As Boolean
Public gB_findGpuMbist_flag As Boolean
Public gB_findSocMbist_flag As Boolean
Public gB_findPwrPin_flag As Boolean

''''20151109 New Update
Private gm_DCCategory As String
Private gm_DCSelector As String
Private gm_ACCategory As String
Private gm_ACSelector As String
Private gm_TimeSetSheet As String
Private gm_EdgeSetSheet As String
Private gm_LevelsSheet As String

''''Initialize, 20151111 New
Public gB_BIRA_MC000_NRS_flag As Boolean
Public gB_BIRA_MC010_NRS_flag As Boolean
Public gB_BIRA_MC051_NRS_flag As Boolean
Public gB_BIRA_MC051_flag As Boolean
Public gB_BIRA_MC000_flag As Boolean

''''20151224 New for Mbist CharZ
Public gD_corePWR_Value As New PinListData

Public Function auto_parse_MBIST_ChkList(bistType As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_parse_MBIST_ChkList"
    
    Dim sheetName As String
    Dim mysheet As Worksheet
    Dim myCell As Object
    Dim offCell As Object
    Dim myCell_Header As Object
    
    Dim myCellA1 As Object
    Dim m_A1_rowCnt As Long
    
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim n As Long
    Dim m_cellCnt As Long
    Dim m_cellStr As String
    Dim m_offcolStr As String
    Dim m_lastrow As Long
    Dim m_lastNCnt As Long

    Dim find_1stHeader As Boolean
    Dim find_AllHeader As Boolean
    
    Dim idx_PModeName_row As Long
    Dim idx_PModeName As Long
    Dim idx_Patt_PMode As Long
    Dim idx_BlockName As Long
    Dim idx_Patt_Block As Long
    
    Dim idx_GroupName As Long
    Dim idx_PMode As Long
    Dim idx_Block As Long
    Dim idx_Debug As Long
    Dim idx_ChangeFlag As Long

    Dim idx_VoltSet As Long
    Dim idx_DCCate As Long
    Dim idx_DCSele As Long
    Dim idx_ACCate As Long
    Dim idx_ACSele As Long
    
    Dim idx_PwrPin As Long
    Dim idx_VStart As Long
    Dim idx_VStop As Long
    Dim idx_VStep As Long
    Dim idx_VEnable As Long

    Dim idx_END As Long

    Dim m_cellArr() As String
    Dim m_cellArr_New() As String
    Dim m_cellArr_elem As String
    Dim m_cellArr_elem_subArr() As String
    Dim elemCNT As Long
    Dim m_cnt As Long
    Dim m_idx As Long
    Dim m_pattname As String
    Dim m_pattRawname As String
    Dim m_debug_PMode As String
    Dim m_debug_Block As String
    
    Dim m_tmpStr As String
    Dim m_elem_ST As Long
    Dim m_elem_SP As Long
    Dim m_GroupCnt As Long
    Dim m_PModePattCnt As Long
    Dim m_BlockPattCnt As Long
    Dim m_VoltSetCnt As Long
    Dim m_PwrPinCnt As Long
    Dim m_status As Boolean
    Dim m_patArr() As String
    Dim m_patcount As Long

    ''''-------------------------------
    DebugPrtImm = False
    DebugPrtDlog = False
    ''''-------------------------------
    
    bistType = UCase(bistType)
    
    gB_findPwrPin_flag = False ''''Initial

    If (bistType = "CPU") Then
        gB_findCpuMbist_flag = False
        sheetName = gS_CpuMbist_sheetName
    ElseIf (bistType = "GPU") Then
        gB_findGpuMbist_flag = False
        sheetName = gS_GpuMbist_sheetName
    ElseIf (bistType = "SOC") Then
        gB_findSocMbist_flag = False
        sheetName = gS_SocMbist_sheetName
    End If

    Set mysheet = Sheets(sheetName)
    
    Set myCellA1 = mysheet.Range("A1")
    m_A1_rowCnt = 0
    
    Set myCell = mysheet.Range("A1")
    m_cellStr = UCase(Trim(myCell.Value))
    
    ''DebugPrintLog "Row=" & myCell.Row & ", Column=" & myCellA1.Column & ", End_Column=" & myCellA1.End(xlToRight).Column
    ''DebugPrintLog "A1...Row=" & myCell.Row & ", Column=" & myCell.Column & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"
    
    find_1stHeader = False
    find_AllHeader = False
    m_A1_rowCnt = 0
    m_GroupCnt = 0
    m_PModePattCnt = 0
    m_BlockPattCnt = 0
    m_VoltSetCnt = 0

    ''''At First, finding all Header Index to identify every column.
    ''''Here 'Do...Loop While()' is to search from Up to Down
    Do
        DebugPrintLog "0...Row=" & myCell.Row & ", Column=" & myCell.Column & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"

        ''''1......To find the 1st Word "Performance"
        ''''Here Search Cell from Left to Right
        If (find_1stHeader = False) Then
            m_cellCnt = 0
            Do While (m_cellCnt < 11)
                ''DebugPrintLog "1...Row=" & myCell.Row & ", Column=" & myCell.Column & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"
    
                If (m_cellStr Like UCase("Performance*Name")) Then
                    idx_PModeName_row = myCell.Row
                    idx_PModeName = myCell.Column
                    
                    DebugPrintLog "1...(Performance Name) Row=" & idx_PModeName_row & ", Column=" & idx_PModeName & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"
                    find_1stHeader = True
                    Exit Do
                End If
    
                ''''if cell search from left to right, (rowOffset:=0, columnOffset:=1)
                Set myCell = myCell.Offset(rowOffset:=0, columnOffset:=1)
                m_cellStr = UCase(Trim(myCell.Value))
                m_cellCnt = m_cellCnt + 1
            Loop
        End If

        ''''2......To find the following Header Words
        ''''By each Header, get the related parameters.
        If (find_1stHeader) Then
            ''''<MUST> Point to the 1st cell of the row which contains the 1st Header Word
            Set myCell = myCellA1.Offset(rowOffset:=idx_PModeName_row - 1, columnOffset:=0)
            m_cellStr = UCase(Trim(myCell.Value))
            DebugPrintLog "2...(find_1stHeader=True) Row=" & myCell.Row & ", Column=" & myCell.Column & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"

            m = 0
            
            Do While (m_offcolStr <> "END")
                ''''if cell search from left to right, (rowOffset:=0, columnOffset:=1)
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=m) ''''search cell from Left to Right
                m_offcolStr = UCase(Trim(offCell.Value))
                m = m + 1
                ''''-------------------------
                ''''Column Sequence
                ''''-------------------------
                ''''Performance Name
                ''''Pattern_PMode
                ''''Block Name
                ''''Pattern_Block
                ''''
                ''''GroupName
                ''''Performance
                ''''Block
                ''''Debug
                ''''ChangeFlag
                ''''
                ''''Voltage Setting
                ''''DCCategory
                ''''DCSelector
                ''''ACCategory
                ''''ACSelector
                ''''
                ''''Power_Pin
                ''''V_Start
                ''''V_Stop
                ''''V_Step
                ''''Loop_Enable
                ''''
                ''''End
                ''''-------------------------
                
                If (m_offcolStr Like UCase("Performance*Name")) Then
                    idx_PModeName = m
                
                ElseIf (m_offcolStr Like UCase("Pattern*PMode")) Then
                    idx_Patt_PMode = m
                
                ElseIf (m_offcolStr Like UCase("Block*Name")) Then
                    idx_BlockName = m
                
                ElseIf (m_offcolStr Like UCase("Pattern*Block")) Then
                    idx_Patt_Block = m

                ElseIf (m_offcolStr = UCase("GroupName")) Then
                    idx_GroupName = m
                
                ElseIf (m_offcolStr = UCase("Performance")) Then
                    idx_PMode = m
                    
                ElseIf (m_offcolStr = UCase("Block")) Then
                    idx_Block = m
                
                ElseIf (m_offcolStr = UCase("Debug")) Then
                    idx_Debug = m
                
                ElseIf (m_offcolStr = UCase("ChangeFlag")) Then
                    idx_ChangeFlag = m
                
                ElseIf (m_offcolStr Like UCase("Voltage*Setting")) Then
                    idx_VoltSet = m
                
                ElseIf (m_offcolStr = UCase("DCCategory")) Then
                    idx_DCCate = m
                
                ElseIf (m_offcolStr = UCase("DCSelector")) Then
                    idx_DCSele = m
                
                ElseIf (m_offcolStr = UCase("ACCategory")) Then
                    idx_ACCate = m
                    
                ElseIf (m_offcolStr = UCase("ACSelector")) Then
                    idx_ACSele = m

                ElseIf (m_offcolStr = UCase("Power_Pin")) Then
                    idx_PwrPin = m
                
                ElseIf (m_offcolStr = UCase("V_Start")) Then
                    idx_VStart = m
                
                ElseIf (m_offcolStr = UCase("V_Stop")) Then
                    idx_VStop = m
                
                ElseIf (m_offcolStr = UCase("V_Step")) Then
                    idx_VStep = m
                
                ElseIf (m_offcolStr = UCase("Loop_Enable")) Then
                    idx_VEnable = m

                End If
                
                DebugPrintLog "3...input m_offcolStr=" + m_offcolStr + " (" & offCell.Value & "), Index(Column)=" + CStr(m)
                
            Loop ''''end of Do While (m_offcolStr <> "END")
            
            If (m_offcolStr = ("END")) Then
                idx_END = m
                find_AllHeader = True
            End If
        End If

        ''''if cell search from up   to down,  (rowOffset:=1, columnOffset:=0)
        ''''if cell search from left to right, (rowOffset:=0, columnOffset:=1)
        ''''Here it MUST be use A1 as reference cell
        m_A1_rowCnt = m_A1_rowCnt + 1
        Set myCell = myCellA1.Offset(rowOffset:=m_A1_rowCnt, columnOffset:=0)
        m_cellStr = UCase(Trim(myCell.Value))
    Loop While (find_AllHeader = False)
    
    ''''After getting all Headers, process/get every content inside.
    If (find_AllHeader) Then
    
        ''''<MUST> Point to the 1st cell of the row which contains the 1st Header Word
        Set myCell_Header = myCellA1.Offset(rowOffset:=idx_PModeName_row - 1, columnOffset:=0)
        m_cellStr = UCase(Trim(myCell_Header.Value))
        DebugPrintLog "4...(find_AllHeader=True) Row=" & myCell_Header.Row & ", Column=" & myCell_Header.Column & ", Cell=" & myCell_Header.Value & " (m_cellStr=" + m_cellStr + ")"
        
        ''''initialize -----------------------------------------
        If (bistType = "CPU") Then
            ReDim CpuMbist.Category(100)
            ReDim CpuMbist_PMode.Category(100)
            ReDim CpuMbist_Block.Category(100)
            ReDim CpuMbist_VoltSet.Category(100)
            ReDim CpuMbist_Power.Category(100)
        ElseIf (bistType = "GPU") Then
        ElseIf (bistType = "SOC") Then
        End If
        ''''----------------------------------------------------
        
        ''''Then get the following parameter per Header
        m = 0
        Do While (m <= idx_END)
            m = m + 1 ''''Column direction
            n = 0 ''''index and row direction

            Set myCell = myCell_Header.Offset(rowOffset:=0, columnOffset:=(m - 1)) ''''rowOffset MUST be always '0'
            m_cellStr = UCase(Trim(myCell.Value))
            m_lastrow = myCell.End(xlDown).Row
            m_lastNCnt = m_lastrow - idx_PModeName_row
            DebugPrintLog "5...input Header =" + m_cellStr + " (" & myCell.Value & "), m=" + CStr(m) + ", LastRow=" + CStr(m_lastrow) + ", LastNCnt=" + CStr(m_lastNCnt)
            
            Select Case (m)
            Case idx_END
                If (bistType = "CPU") Then
                    gB_findCpuMbist_flag = True
                ElseIf (bistType = "GPU") Then
                    gB_findGpuMbist_flag = True
                ElseIf (bistType = "SOC") Then
                    gB_findSocMbist_flag = True
                End If
                Exit Do ''''end

            Case idx_PModeName
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", Category Index n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_PMode.Category(n).Name = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                m_PModePattCnt = n
                If (bistType = "CPU") Then
                    ReDim Preserve CpuMbist_PMode.Category(n - 1)
                ElseIf (bistType = "GPU") Then
                ElseIf (bistType = "SOC") Then
                End If
                n = 0
                
            Case idx_Patt_PMode
                ''Do While (n < m_lastNCnt)
                Do While (n < m_PModePattCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_PMode.Category(n).PattName = m_cellStr
                        ''''------------------------------------------------------------------------
                        m_status = GetPatListFromPatternSet(m_cellStr, m_patArr, m_patcount)
                        If (m_status) Then
                            If (m_patcount = 1) Then
                                CpuMbist_PMode.Category(n).PattRawName = m_patArr(0)
                            Else
                                ''''there are multi-patterns in the PattSet, it needs another process later on
                                CpuMbist_PMode.Category(n).PattRawName = m_cellStr
                            End If
                        Else
                            CpuMbist_PMode.Category(n).PattRawName = "NA"
                        End If
                        ''''------------------------------------------------------------------------
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_BlockName
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_Block.Category(n).Name = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                m_BlockPattCnt = n
                If (bistType = "CPU") Then
                    ReDim Preserve CpuMbist_Block.Category(n - 1)
                ElseIf (bistType = "GPU") Then
                ElseIf (bistType = "SOC") Then
                End If
                n = 0

            Case idx_Patt_Block
                ''Do While (n < m_lastNCnt)
                Do While (n < m_BlockPattCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_Block.Category(n).PattName = m_cellStr
                        ''''------------------------------------------------------------------------
                        m_status = GetPatListFromPatternSet(m_cellStr, m_patArr, m_patcount)
                        If (m_status) Then
                            If (m_patcount = 1) Then
                                CpuMbist_Block.Category(n).PattRawName = m_patArr(0)
                            Else
                                ''''there are multi-patterns in the PattSet, it needs another process later on
                                CpuMbist_Block.Category(n).PattRawName = m_cellStr
                            End If
                        Else
                            CpuMbist_Block.Category(n).PattRawName = "NA"
                        End If
                        ''''------------------------------------------------------------------------
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_GroupName
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist.Category(n).Name = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                m_GroupCnt = n
                If (bistType = "CPU") Then
                    ReDim Preserve CpuMbist.Category(n - 1)
                ElseIf (bistType = "GPU") Then
                ElseIf (bistType = "SOC") Then
                End If
                n = 0

            Case idx_PMode
                ''Do While (n < m_lastNCnt)
                Do While (n < m_GroupCnt)
                    m_cnt = 0 ''''is used to count how many elements of XXX.PMode() including empty,NA,N/A
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n

                    If (bistType = "CPU") Then
                        CpuMbist.Category(n).PMode_RAW = m_cellStr
                        ReDim CpuMbist.Category(n).PMode(100) ''''Initialize
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If

                    If (m_cellStr <> "" And UCase(m_cellStr) <> "NA" And UCase(m_cellStr) <> "N/A") Then
                        ''''------------------------------------------------
                        ''''Split m_cellStr as String array
                        ''''------------------------------------------------
                        m_cellArr = Split(Trim(m_cellStr), ",")
                        elemCNT = 0
                        For j = 0 To UBound(m_cellArr)
                            m_cellArr_elem = Trim(m_cellArr(j))
                            If (m_cellArr_elem <> "") Then
                                m_cellArr(elemCNT) = m_cellArr_elem
                                
                                If (bistType = "CPU") Then
                                    CpuMbist.Category(n).PMode(j).Name = m_cellArr_elem
                                    ''''get the related patter name
                                    m_idx = CpuMbist_PMode_Index(m_cellArr_elem)
                                    m_pattname = CpuMbist_PMode.Category(m_idx).PattName
                                    m_pattRawname = CpuMbist_PMode.Category(m_idx).PattRawName
                                    CpuMbist.Category(n).PMode(j).PattName = m_pattname
                                    CpuMbist.Category(n).PMode(j).PattRawName = m_pattRawname
                                ElseIf (bistType = "GPU") Then
                                ElseIf (bistType = "SOC") Then
                                End If
                                elemCNT = elemCNT + 1
                                m_cnt = m_cnt + 1
                            End If
                        Next j
                        If (elemCNT >= 1) Then
                            ReDim Preserve m_cellArr(elemCNT - 1)
                        End If
                        ''''------------------------------------------------
                    Else
                        If (bistType = "CPU") Then
                            CpuMbist.Category(n).PMode(m_cnt).Name = "NA"
                            CpuMbist.Category(n).PMode(m_cnt).PattName = "NA"
                            CpuMbist.Category(n).PMode(m_cnt).PattRawName = "NA"
                        ElseIf (bistType = "GPU") Then
                        ElseIf (bistType = "SOC") Then
                        End If
                        m_cnt = m_cnt + 1
                    End If

                    ''''Finalize the dimension
                    If (bistType = "CPU") Then
                        ReDim Preserve CpuMbist.Category(n).PMode(m_cnt - 1) ''''MUST be (m_cnt-1)
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    
                    n = n + 1
                Loop
                n = 0

            Case idx_Block
                ''Do While (n < m_lastNCnt)
                Do While (n < m_GroupCnt)
                    m_cnt = 0 ''''is used to count how many elements of XXX.Block() including empty,NA,N/A
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n

                    If (bistType = "CPU") Then
                        CpuMbist.Category(n).Block_RAW = m_cellStr
                        ReDim CpuMbist.Category(n).Block(200) ''''Initialize
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If

                    If (m_cellStr <> "" And UCase(m_cellStr) <> "NA" And UCase(m_cellStr) <> "N/A") Then
                        ''''------------------------------------------------
                        ''''Split m_cellStr as String array
                        ''''------------------------------------------------
                        m_cellArr = Split(Trim(m_cellStr), ",")
                        ReDim Preserve m_cellArr(UBound(m_cellArr))
                        ReDim m_cellArr_New(200) ''''Initialize
                        
                        elemCNT = 0
                        For j = 0 To UBound(m_cellArr)
                            m_cellArr_elem = Trim(m_cellArr(j))
                            If (m_cellArr_elem <> "") Then
                                m_tmpStr = UCase(m_cellArr_elem)
                                If (m_tmpStr Like UCase("B*_B*")) Then
                                    ReDim m_cellArr_elem_subArr(1) ''''initialize
                                    m_tmpStr = Replace(m_tmpStr, "B", "") ''''remove 'B' characters
                                    m_cellArr_elem_subArr = Split(m_tmpStr, "_")
                                
                                    If (IsNumeric(m_cellArr_elem_subArr(0)) And IsNumeric(m_cellArr_elem_subArr(1))) Then
                                        m_elem_ST = CLng(m_cellArr_elem_subArr(0))
                                        m_elem_SP = CLng(m_cellArr_elem_subArr(1))
    
                                        For i = m_elem_ST To m_elem_SP
                                            m_cellArr_New(elemCNT) = "B" + Format(i, "00")
                                            elemCNT = elemCNT + 1
                                        Next i
                                        elemCNT = elemCNT - 1 ''''check
                                    Else
                                        TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out. (Block element = " + m_cellArr_elem + ")"
                                        GoTo errHandler
                                    End If
                                Else
                                    ''''Only one block mode case
                                    m_cellArr_New(elemCNT) = m_cellArr_elem
                                End If
                                elemCNT = elemCNT + 1
                            End If ''''end of If (m_cellArr_elem <> "") Then
                        Next j
                        If (elemCNT >= 1) Then
                            ReDim Preserve m_cellArr_New(elemCNT - 1)
    
                            If (bistType = "CPU") Then
                                For i = 0 To UBound(m_cellArr_New)
                                    m_cellArr_elem = m_cellArr_New(i)
                                    CpuMbist.Category(n).Block(i).Name = m_cellArr_elem
                                    ''''get the related patter name
                                    m_idx = CpuMbist_Block_Index(m_cellArr_elem)
                                    m_pattname = CpuMbist_Block.Category(m_idx).PattName
                                    m_pattRawname = CpuMbist_Block.Category(m_idx).PattRawName
                                    CpuMbist.Category(n).Block(i).PattName = m_pattname
                                    CpuMbist.Category(n).Block(i).PattRawName = m_pattRawname
                                    m_cnt = m_cnt + 1
                                Next i
                            ElseIf (bistType = "GPU") Then
                            ElseIf (bistType = "SOC") Then
                            End If
                            
                        End If
                        ''''------------------------------------------------
                    Else
                        ''''case:: Empty or 'NA' / 'N/A'
                        If (bistType = "CPU") Then
                            CpuMbist.Category(n).Block(m_cnt).Name = "NA"
                            CpuMbist.Category(n).Block(m_cnt).PattName = "NA"
                            CpuMbist.Category(n).Block(m_cnt).PattRawName = "NA"
                        ElseIf (bistType = "GPU") Then
                        ElseIf (bistType = "SOC") Then
                        End If
                        m_cnt = m_cnt + 1
                    End If
                    
                    ''''Finalize the dimension
                    If (bistType = "CPU") Then
                        ReDim Preserve CpuMbist.Category(n).Block(m_cnt - 1) ''''MUST (m_cnt -1)
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If

                    n = n + 1
                Loop
                n = 0

            Case idx_Debug
                ''Do While (n < m_lastNCnt)
                Do While (n < m_GroupCnt)
                    m_cnt = 0 ''''is used to count how many elements of XXX.DebugMode() including empty,NA,N/A
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    
                    If (bistType = "CPU") Then
                        ReDim CpuMbist.Category(n).DebugMode(100) ''''Initialize
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    
                    If (m_cellStr <> "" And UCase(m_cellStr) <> "NA" And UCase(m_cellStr) <> "N/A") Then
                        ''''------------------------------------------------
                        ''''Split m_cellStr as String array
                        ''''------------------------------------------------
                        m_cellArr = Split(Trim(m_cellStr), ",")
                        elemCNT = 0
                        For j = 0 To UBound(m_cellArr)
                            m_cellArr_elem = Trim(m_cellArr(j))
                            If (m_cellArr_elem <> "") Then
                                m_cellArr(elemCNT) = m_cellArr_elem
                                ReDim m_cellArr_elem_subArr(1) ''''initialize
                                m_cellArr_elem_subArr = Split(m_cellArr_elem, "_")
                                ''ReDim Preserve m_cellArr_elem_subArr(UBound(m_cellArr_elem_subArr))
                                m_debug_PMode = m_cellArr_elem_subArr(0)
                                m_debug_Block = m_cellArr_elem_subArr(1)
                                
                                If (bistType = "CPU") Then
                                    CpuMbist.Category(n).DebugMode(j).Name = m_cellArr_elem
                                    ''''get the related patter name
                                    m_idx = CpuMbist_PMode_Index(m_debug_PMode)
                                    m_pattname = CpuMbist_PMode.Category(m_idx).PattName
                                    m_pattRawname = CpuMbist_PMode.Category(m_idx).PattRawName
                                    CpuMbist.Category(n).DebugMode(j).PMode = m_debug_PMode
                                    CpuMbist.Category(n).DebugMode(j).PMode_PattName = m_pattname
                                    CpuMbist.Category(n).DebugMode(j).PMode_PattRawName = m_pattRawname
                                    ''''get the related patter name
                                    m_idx = CpuMbist_Block_Index(m_debug_Block)
                                    m_pattname = CpuMbist_Block.Category(m_idx).PattName
                                    m_pattRawname = CpuMbist_Block.Category(m_idx).PattRawName
                                    CpuMbist.Category(n).DebugMode(j).Block = m_debug_Block
                                    CpuMbist.Category(n).DebugMode(j).Block_PattName = m_pattname
                                    CpuMbist.Category(n).DebugMode(j).Block_PattRawName = m_pattRawname
                                ElseIf (bistType = "GPU") Then
                                ElseIf (bistType = "SOC") Then
                                End If
                                elemCNT = elemCNT + 1
                                m_cnt = m_cnt + 1
                            End If
                        Next j
                        If (elemCNT >= 1) Then
                            ReDim Preserve m_cellArr(elemCNT - 1)
                        End If
                        ''''------------------------------------------------
                    Else
                        ''''case:: Empty or 'NA' / 'N/A'
                        If (bistType = "CPU") Then
                            CpuMbist.Category(n).DebugMode(m_cnt).Name = "NA"
                            CpuMbist.Category(n).DebugMode(m_cnt).PMode = "NA"
                            CpuMbist.Category(n).DebugMode(m_cnt).PMode_PattName = "NA"
                            CpuMbist.Category(n).DebugMode(m_cnt).PMode_PattRawName = "NA"
                            CpuMbist.Category(n).DebugMode(m_cnt).Block = "NA"
                            CpuMbist.Category(n).DebugMode(m_cnt).Block_PattName = "NA"
                            CpuMbist.Category(n).DebugMode(m_cnt).Block_PattRawName = "NA"
                        ElseIf (bistType = "GPU") Then
                        ElseIf (bistType = "SOC") Then
                        End If
                        m_cnt = m_cnt + 1
                    End If

                    ''''Finalize the dimension
                    If (bistType = "CPU") Then
                        ReDim Preserve CpuMbist.Category(n).DebugMode(m_cnt - 1)
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_ChangeFlag
                ''Do While (n < m_lastNCnt)
                Do While (n < m_GroupCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist.Category(n).ChangeFlag = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_VoltSet
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_VoltSet.Category(n).Name = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                m_VoltSetCnt = n
                If (bistType = "CPU") Then
                    ReDim Preserve CpuMbist_VoltSet.Category(n - 1)
                ElseIf (bistType = "GPU") Then
                ElseIf (bistType = "SOC") Then
                End If
                n = 0

            Case idx_DCCate
                ''Do While (n < m_lastNCnt)
                Do While (n < m_VoltSetCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_VoltSet.Category(n).DCCate = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_DCSele
                ''Do While (n < m_lastNCnt)
                Do While (n < m_VoltSetCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_VoltSet.Category(n).DCSele = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_ACCate
                ''Do While (n < m_lastNCnt)
                Do While (n < m_VoltSetCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_VoltSet.Category(n).ACCate = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_ACSele
                ''Do While (n < m_lastNCnt)
                Do While (n < m_VoltSetCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_VoltSet.Category(n).ACSele = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_PwrPin
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_Power.Category(n).pwrpin = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                m_PwrPinCnt = n
                gB_findPwrPin_flag = True
                If (bistType = "CPU") Then
                    ReDim Preserve CpuMbist_Power.Category(n - 1)
                ElseIf (bistType = "GPU") Then
                ElseIf (bistType = "SOC") Then
                End If
                n = 0

            Case idx_VStart
                ''Do While (n < m_lastNCnt)
                Do While (n < m_PwrPinCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_Power.Category(n).V_Start = CDbl(m_cellStr)
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_VStop
                ''Do While (n < m_lastNCnt)
                Do While (n < m_PwrPinCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_Power.Category(n).V_Stop = CDbl(m_cellStr)
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_VStep
                ''Do While (n < m_lastNCnt)
                Do While (n < m_PwrPinCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_Power.Category(n).V_Step = CDbl(m_cellStr)
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_VEnable
                ''Do While (n < m_lastNCnt)
                Do While (n < m_PwrPinCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    If (m_cellStr = "Y" Or m_cellStr = "YES") Then
                        m_cellStr = "Y"
                    Else
                        ''''case "", "No", "NA", "N/A"
                        m_cellStr = "N"
                    End If
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.Row & ", n=" & n
                    If (bistType = "CPU") Then
                        CpuMbist_Power.Category(n).Enable = m_cellStr
                    ElseIf (bistType = "GPU") Then
                    ElseIf (bistType = "SOC") Then
                    End If
                    n = n + 1
                Loop
                n = 0

            Case Else
                DebugPrintLog "6...Empty Column(" & m & ") !!!"
            End Select

        Loop ''''end of Do While (m <= idx_END)
        
    End If ''''end of If (find_AllHeader) Then
    
 
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

''''For CpuMbist_PMode.Category()
Public Function CpuMbist_PMode_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CpuMbist_PMode_Index"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
    For i = 0 To UBound(CpuMbist_PMode.Category)
        If (UCase(myStr) = UCase(CpuMbist_PMode.Category(i).Name)) Then
            CpuMbist_PMode_Index = i
            match_Flag = True
            Exit For
        End If
    Next i

    If (match_Flag = False) Then
        CpuMbist_PMode_Index = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

''''For CpuMbist_Block.Category()
Public Function CpuMbist_Block_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CpuMbist_Block_Index"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
    For i = 0 To UBound(CpuMbist_Block.Category)
        If (UCase(myStr) = UCase(CpuMbist_Block.Category(i).Name)) Then
            CpuMbist_Block_Index = i
            match_Flag = True
            Exit For
        End If
    Next i

    If (match_Flag = False) Then
        CpuMbist_Block_Index = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''For CpuMbist.Category()
Public Function CpuMbist_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CpuMbist_Index"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
    For i = 0 To UBound(CpuMbist.Category)
        If (UCase(myStr) = UCase(CpuMbist.Category(i).Name)) Then
            CpuMbist_Index = i
            match_Flag = True
            Exit For
        End If
    Next i

    If (match_Flag = False) Then
        CpuMbist_Index = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''For CpuMbist.Category(index).PMode()
Public Function CpuMbist_subPMode_Index(Index As Long, myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CpuMbist_subPMode_Index"

    Dim j As Long
    Dim match_Flag As Boolean

    match_Flag = False
    
    For j = 0 To UBound(CpuMbist.Category(Index).PMode)
        If (UCase(myStr) = UCase(CpuMbist.Category(Index).PMode(j).Name)) Then
            CpuMbist_subPMode_Index = j
            match_Flag = True
            Exit For
        End If
    Next j

    If (match_Flag = False) Then
        CpuMbist_subPMode_Index = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''For CpuMbist.Category(index).Block()
Public Function CpuMbist_subBlock_Index(Index As Long, myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CpuMbist_subBlock_Index"

    Dim j As Long
    Dim match_Flag As Boolean

    match_Flag = False
    
    For j = 0 To UBound(CpuMbist.Category(Index).Block)
        If (UCase(myStr) = UCase(CpuMbist.Category(Index).Block(j).Name)) Then
            CpuMbist_subBlock_Index = j
            match_Flag = True
            Exit For
        End If
    Next j

    If (match_Flag = False) Then
        CpuMbist_subBlock_Index = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''For CpuMbist.Category(index).DebugMode()
Public Function CpuMbist_subDebugMode_Index(Index As Long, myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CpuMbist_subDebugMode_Index"

    Dim j As Long
    Dim match_Flag As Boolean

    match_Flag = False
    
    For j = 0 To UBound(CpuMbist.Category(Index).DebugMode)
        If (UCase(myStr) = UCase(CpuMbist.Category(Index).DebugMode(j).Name)) Then
            CpuMbist_subDebugMode_Index = j
            match_Flag = True
            Exit For
        End If
    Next j

    If (match_Flag = False) Then
        CpuMbist_subDebugMode_Index = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''For CpuMbist_VoltSet.Category()
Public Function CpuMbist_VoltSet_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CpuMbist_VoltSet_Index"

    Dim j As Long
    Dim match_Flag As Boolean

    match_Flag = False
    
    For j = 0 To UBound(CpuMbist_VoltSet.Category)
        If (UCase(myStr) = UCase(CpuMbist_VoltSet.Category(j).Name)) Then
            CpuMbist_VoltSet_Index = j
            match_Flag = True
            Exit For
        End If
    Next j

    If (match_Flag = False) Then
        CpuMbist_VoltSet_Index = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''For CpuMbist_Power.Category()
Public Function CpuMbist_Power_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CpuMbist_Power_Index"

    Dim j As Long
    Dim match_Flag As Boolean

    match_Flag = False
    
    For j = 0 To UBound(CpuMbist_Power.Category)
        If (UCase(myStr) = UCase(CpuMbist_Power.Category(j).Name)) Then
            CpuMbist_Power_Index = j
            match_Flag = True
            Exit For
        End If
    Next j

    If (match_Flag = False) Then
        CpuMbist_Power_Index = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_Mbist_GetLoopIndex_DM_PM_BM(ByRef m_LoopIndex_DM As Long, _
                                                 ByRef m_LoopIndex_PM As Long, _
                                                 ByRef m_LoopIndex_BM As Long, _
                                                 Optional showPrint As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Mbist_GetLoopIndex_DM_PM_BM"
    
    Dim m_site As Variant
    Dim m_dlgStr As String
    
    Dim m_LCount_DM_siteVar As String
    Dim m_LCount_PM_siteVar As String
    Dim m_LCount_BM_siteVar As String
    
    Dim m_LCount_DM As Long
    Dim m_LCount_PM As Long
    Dim m_LCount_BM As Long
    
    Dim m_LP_DM_siteVar As String
    Dim m_LP_PM_siteVar As String
    Dim m_LP_BM_siteVar As String

    m_LCount_DM_siteVar = "LCount_DM"
    m_LCount_PM_siteVar = "LCount_PM"
    m_LCount_BM_siteVar = "LCount_BM"

    m_LP_DM_siteVar = "LP_DM"
    m_LP_PM_siteVar = "LP_PM"
    m_LP_BM_siteVar = "LP_BM"

    For Each m_site In TheExec.Sites
        m_LCount_DM = TheExec.Sites(m_site).SiteVariableValue(m_LCount_DM_siteVar)
        m_LCount_PM = TheExec.Sites(m_site).SiteVariableValue(m_LCount_PM_siteVar)
        m_LCount_BM = TheExec.Sites(m_site).SiteVariableValue(m_LCount_BM_siteVar)
    
        m_LoopIndex_DM = TheExec.Sites(m_site).SiteVariableValue(m_LP_DM_siteVar)
        m_LoopIndex_PM = TheExec.Sites(m_site).SiteVariableValue(m_LP_PM_siteVar)
        m_LoopIndex_BM = TheExec.Sites(m_site).SiteVariableValue(m_LP_BM_siteVar)
        Exit For ''''because all sites have the same
    Next m_site

    ''''<Important> 20151106 update to fix Debug Mode
    If (m_LoopIndex_BM >= m_LCount_BM) Then m_LoopIndex_BM = 0 ''''<MUST> for the next PM loop
    If (m_LoopIndex_PM >= m_LCount_PM) Then m_LoopIndex_PM = 0 ''''<MUST> for the next DM loop
    If (m_LoopIndex_DM > m_LCount_DM) Then m_LoopIndex_DM = 0  ''''<MUST> for the next for-loop (DM-PM-BM) block

    If (showPrint) Then
        PrintDataLog ""
        m_dlgStr = ""
        m_dlgStr = FormatNumeric("GroupName = " + gS_Mbist_GroupName, 15 + (Len(gS_Mbist_GroupName)))
        m_dlgStr = m_dlgStr + ", Current LoopIndex "
        m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_LP_DM_siteVar, 6) + " = " + FormatNumeric(m_LoopIndex_DM, -3)
        m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_LP_PM_siteVar, 6) + " = " + FormatNumeric(m_LoopIndex_PM, -3)
        m_dlgStr = m_dlgStr + ", " + FormatNumeric(m_LP_BM_siteVar, 6) + " = " + FormatNumeric(m_LoopIndex_BM, -3)
        PrintDataLog m_dlgStr
    End If
    
    auto_Mbist_GetLoopIndex_DM_PM_BM = 1
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_Mbist_GetBinFlagName(m_ChangeFlag As String, _
                                          m_instNameArr() As String, _
                                          m_headStr As String, _
                                          m_lastStr As String, _
                                          m_PMname As String, _
                                          m_BMname As String, _
                                          m_testName As String) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Mbist_GetBinFlagName"

    Dim i As Long
    Dim m_TestInstName As String
    Dim m_BinFlagName As String
    Dim m_elemStr As String
    Dim m_d64name As String
    Dim m_PMnameX As String ''''with underscore '_'
    Dim m_BMnameX As String ''''with underscore '_'
    Dim m_PMBMname As String
    Dim m_XXXXname As String ''''as 'harddefect', 'bitcell', 'D64_harddefect', 'D64_bitcell'
    Dim m_middname As String '''Middle name string
    
    m_TestInstName = UCase(Trim(TheExec.DataManager.InstanceName))
    
    If (Trim(m_PMname) <> "") Then
        m_PMnameX = "_" + m_PMname
    Else
        m_PMnameX = ""
    End If
    
    If (Trim(m_BMname) <> "") Then
        m_BMnameX = "_" + m_BMname
    Else
        m_BMnameX = ""
    End If
    
    ''''20151120 update
    ''''------------------------------------------------------------------------
    If (Trim(m_PMname) <> "" And Trim(m_BMname) <> "") Then
        m_PMBMname = m_PMnameX + m_BMnameX
    ElseIf (Trim(m_PMname) <> "" And Trim(m_BMname) = "") Then
        m_PMBMname = m_PMnameX
    ElseIf (Trim(m_PMname) = "" And Trim(m_BMname) <> "") Then
        m_PMBMname = m_BMnameX
    Else
        m_PMBMname = ""
    End If
    
    If (UCase(gS_Mbist_GroupName) Like UCase("*D64*harddefect*")) Then
        m_d64name = "_d64"
        m_XXXXname = "_harddefect"
    ElseIf (UCase(gS_Mbist_GroupName) Like UCase("*D64*bitcell*")) Then
        m_d64name = "_d64"
        m_XXXXname = "_bitcell"
    ElseIf (UCase(gS_Mbist_GroupName) Like UCase("*harddefect*")) Then
        m_d64name = ""
        m_XXXXname = "_harddefect"
    ElseIf (UCase(gS_Mbist_GroupName) Like UCase("*bitcell*")) Then
        m_d64name = ""
        m_XXXXname = "_bitcell"
    ElseIf (UCase(gS_Mbist_GroupName) Like UCase("*D64*")) Then ''''Non HardDefect and BitCell but D64
        m_d64name = "_d64"
        m_XXXXname = ""
    Else
        m_d64name = ""
        m_XXXXname = ""
    End If
    
    If (Trim(m_d64name) <> "") Then
        If (Trim(m_XXXXname) <> "" And Trim(m_PMBMname) <> "") Then
            m_middname = m_d64name + m_XXXXname + m_PMBMname

        ElseIf (Trim(m_XXXXname) = "" And Trim(m_PMBMname) <> "") Then
            m_middname = m_d64name + m_PMBMname

        ElseIf (Trim(m_XXXXname) <> "" And Trim(m_PMBMname) = "") Then
            m_middname = m_d64name + m_XXXXname

        Else
            m_middname = m_d64name
        End If
    Else
        ''''non D64 case
        If (Trim(m_XXXXname) <> "" And Trim(m_PMBMname) <> "") Then
            m_middname = m_XXXXname + m_PMBMname

        ElseIf (Trim(m_XXXXname) = "" And Trim(m_PMBMname) <> "") Then
            m_middname = m_PMBMname

        ElseIf (Trim(m_XXXXname) <> "" And Trim(m_PMBMname) = "") Then
            m_middname = m_XXXXname

        Else
            m_middname = ""
        End If
    
    End If
    ''''------------------------------------------------------------------------
    For i = 0 To UBound(m_instNameArr)
        m_elemStr = UCase(m_instNameArr(i))
        If (m_elemStr Like UCase("*Efuse*")) Then
            m_ChangeFlag = "F"
            m_BinFlagName = "Default"
            Exit For
        End If
    Next i
    ''''------------------------------------------------------------------------

    If (m_ChangeFlag = "BT") Then ''''Group=BIRA_ALL case
        If (UCase(gS_Mbist_GroupName) = UCase("BIRA_ALL")) Then
            If (gB_BIRA_MC000_NRS_flag) Then
                m_middname = "_harddefect" + m_PMBMname
                If (UCase(m_TestInstName) Like UCase("*BitCell*")) Then
                    m_middname = "_bitcell" + m_PMBMname
                End If
                
            ElseIf (gB_BIRA_MC010_NRS_flag) Then
                m_middname = "_d64_harddefect" + m_PMBMname
                If (UCase(m_TestInstName) Like UCase("*BitCell*")) Then
                    m_middname = "_d64_bitcell" + m_PMBMname
                End If
                
            ElseIf (gB_BIRA_MC051_NRS_flag) Then
                m_middname = "_MC051" + m_BMnameX
                
            Else
                m_middname = m_PMBMname
            End If
        End If
        m_BinFlagName = "F_" + m_headStr + m_middname + "_" + m_lastStr

        ''''Exception Case
        For i = 0 To UBound(m_instNameArr)
            m_elemStr = UCase(m_instNameArr(i))
''''            If (m_elemStr = UCase("HardDefect")) Then
''''                m_BinFlagName = "Default" ''''use Default from TestFlow
''''                Exit For
''''            ElseIf (m_elemStr = UCase("BitCell")) Then
''''                m_BinFlagName = "Default" ''''use Default from TestFlow
''''                Exit For
''''            End If
            If (m_elemStr Like UCase("*Efuse*")) Then
                m_BinFlagName = "Default" ''''use Default from TestFlow
                Exit For
            End If
        Next i
        
    ElseIf (m_ChangeFlag = "T") Then
        m_BinFlagName = "F_" + m_headStr + m_middname + "_" + m_lastStr
        
    ElseIf (m_ChangeFlag = "RT") Then
        m_BinFlagName = "F_R_" + m_headStr + m_middname + "_" + m_lastStr

    ElseIf (m_ChangeFlag = "B") Then
        m_BinFlagName = "F_" + m_testName
    
    ElseIf (m_ChangeFlag = "F") Then
        m_BinFlagName = "Default" ''''use Default from TestFlow
    Else
        m_BinFlagName = "Default" ''''use Default from TestFlow
    End If
    
    auto_Mbist_GetBinFlagName = m_BinFlagName
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151021, 20151109 Update
''''20151111, Add freqPattSet Parsing
Public Function auto_FuncTest_CpuMbist_Body(ByRef m_pattname As String, _
                                            ByRef m_testName As String, _
                                            ByRef m_BinFlagName As String, _
                                            Optional showPrint As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_FuncTest_CpuMbist_Body"

    ''''CpuMbist---------------------------------
    Dim i As Long
    Dim m_tmpStr As String
    Dim m_grpIndex As Long
    Dim m_grpname As String
    Dim m_grpnameR As String
    Dim m_LoopIndex_DM As Long
    Dim m_LoopIndex_PM As Long
    Dim m_LoopIndex_BM As Long
    
    Dim m_PMname As String
    Dim m_PM_pattname As String
    Dim m_PM_pattRawname As String
    Dim m_BMname As String
    Dim m_BM_pattname As String
    Dim m_BM_pattRawname As String
    Dim m_DMname As String
    Dim m_pattRawname As String
    ''''----------------------------------------
    Dim m_TestInstName As String
    Dim m_TestInstName_Arr() As String
    Dim m_1st_elemStr As String
    Dim m_last_elemStr As String
    Dim m_elemStr As String
    Dim m_elemStrR As String
    ''''----------------------------------------
    Dim m_tailStr As String
    Dim m_VMname As String
    Dim m_DCCate As String
    Dim m_ACCate As String
    ''''----------------------------------------
    ''Dim m_BinFlagName As String
    Dim m_ChangeFlag As String
    ''''----------------------------------------
    
    ''''-------------------------------------------------------------------------------------------------
    '''' Get the related parameters of Mbist
    ''''-------------------------------------------------------------------------------------------------
    m_grpname = gS_Mbist_GroupName
    m_grpIndex = gL_Mbist_GroupName_Index
    Call auto_Mbist_GetLoopIndex_DM_PM_BM(m_LoopIndex_DM, m_LoopIndex_PM, m_LoopIndex_BM, False)

    ''''--------------------------------------------------------------
    ''''<NOTICE>
    ''''--------------------------------------------------------------
    ''''LP_DM For-Loop MUST be from 0 in the Flow and <= LCount_DM
    ''''LP_PM For-Loop MUST be from 0 in the Flow and <  LCount_PM
    ''''LP_BM For-Loop MUST be from 0 in the Flow and <  LCount_BM
    ''''--------------------------------------------------------------
    
    m_DMname = UCase(CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).Name)
    If (m_DMname = "NA") Then
        ''''Case: DebugMode is Disable
        m_PMname = UCase(CpuMbist.Category(m_grpIndex).PMode(m_LoopIndex_PM).Name)
        m_PM_pattname = CpuMbist.Category(m_grpIndex).PMode(m_LoopIndex_PM).PattName
        m_PM_pattRawname = CpuMbist.Category(m_grpIndex).PMode(m_LoopIndex_PM).PattRawName
        
        m_BMname = UCase(CpuMbist.Category(m_grpIndex).Block(m_LoopIndex_BM).Name)
        m_BM_pattname = CpuMbist.Category(m_grpIndex).Block(m_LoopIndex_BM).PattName
        m_BM_pattRawname = CpuMbist.Category(m_grpIndex).Block(m_LoopIndex_BM).PattRawName
    Else
        ''''Case: DebugMode is Enable
        ''''<NOTICE> Only use 'm_LoopIndex_DM' to decide the PMode and Block mode
        m_PMname = UCase(CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).PMode)
        m_PM_pattname = CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).PMode_PattName
        m_PM_pattRawname = CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).PMode_PattRawName
        
        m_BMname = UCase(CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).Block)
        m_BM_pattname = CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).Block_PattName
        m_BM_pattRawname = CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).Block_PattRawName
    End If
    ''''-------------------------------------------------------------------------------------------------
    
    ''''-------------------------------------------------------------------------------------------------
    '''' Set the specified TestName
    ''''-------------------------------------------------------------------------------------------------
    Dim m_idx As Long
    m_grpnameR = Replace(m_grpname, "_", "")
    m_TestInstName = Trim(TheExec.DataManager.InstanceName)
    m_TestInstName_Arr = Split(m_TestInstName, "_")
    ReDim Preserve m_TestInstName_Arr(UBound(m_TestInstName_Arr))
    m_1st_elemStr = UCase(m_TestInstName_Arr(0))
    m_last_elemStr = m_TestInstName_Arr(UBound(m_TestInstName_Arr))
    m_testName = ""
    m_tmpStr = ""

    If (UCase(m_grpname) = UCase("BIRA_ALL")) Then '''' "BIRA_ALL"

        If (UCase(m_TestInstName) Like UCase("CpuPllInit*_MC000_*_NRS_*")) Then
            ''''case PLLInit is MC000_NRS
            gB_BIRA_MC000_NRS_flag = True
            gB_BIRA_MC010_NRS_flag = False
            gB_BIRA_MC051_NRS_flag = False
            gB_BIRA_MC051_flag = False
            gB_BIRA_MC000_flag = False
            m_PMname = "MC000_NRS"

        ElseIf (UCase(m_TestInstName) Like UCase("CpuPllInit*_MC010_*_NRS_*")) Then
            ''''case PLLInit is MC010_NRS
            gB_BIRA_MC000_NRS_flag = False
            gB_BIRA_MC010_NRS_flag = True
            gB_BIRA_MC051_NRS_flag = False
            gB_BIRA_MC051_flag = False
            gB_BIRA_MC000_flag = False
            m_PMname = "MC010_NRS"

        ElseIf (UCase(m_TestInstName) Like UCase("CpuPllInit*_MC051_*_NRS_*")) Then
            ''''case PLLInit is MC051_NRS
            gB_BIRA_MC000_NRS_flag = False
            gB_BIRA_MC010_NRS_flag = False
            gB_BIRA_MC051_NRS_flag = True
            gB_BIRA_MC051_flag = False
            gB_BIRA_MC000_flag = False
            m_PMname = "MC051_NRS"
            
        ElseIf (UCase(m_TestInstName) Like UCase("CpuPllInit*_MC051_*")) Then
            ''''case PLLInit is MC051
            gB_BIRA_MC000_NRS_flag = False
            gB_BIRA_MC010_NRS_flag = False
            gB_BIRA_MC051_NRS_flag = False
            gB_BIRA_MC051_flag = True
            gB_BIRA_MC000_flag = False
            m_PMname = "MC051"

        ElseIf (UCase(m_TestInstName) Like UCase("CpuPllInit*_MC000_*")) Then
            ''''case PLLInit is MC000 (from table)
            gB_BIRA_MC000_NRS_flag = False
            gB_BIRA_MC010_NRS_flag = False
            gB_BIRA_MC051_NRS_flag = False
            gB_BIRA_MC051_flag = False
            gB_BIRA_MC000_flag = True
            m_PMname = m_PMname ''''No Change, follow the table
        
        ElseIf (UCase(m_TestInstName) Like UCase("CpuRepairEFuse*")) Then
            ''''case CpuRepairEFuse*
            gB_BIRA_MC000_NRS_flag = False
            gB_BIRA_MC010_NRS_flag = False
            gB_BIRA_MC051_NRS_flag = False
            gB_BIRA_MC051_flag = False
            gB_BIRA_MC000_flag = True
            m_PMname = m_PMname ''''No Change, follow the table
        End If

        ''''<Important and MUST> It's as reference Index for the following items.
        If (m_1st_elemStr Like UCase("*RepairEFuse*")) Then
            m_PMname = m_PMname ''''No Change, follow the table
        ElseIf (gB_BIRA_MC000_NRS_flag) Then
            m_PMname = "MC000_NRS"
        ElseIf (gB_BIRA_MC010_NRS_flag) Then
            m_PMname = "MC010_NRS"
        ElseIf (gB_BIRA_MC051_NRS_flag) Then
            m_PMname = "MC051_NRS"
        ElseIf (gB_BIRA_MC051_flag) Then
            m_PMname = "MC051"
        ElseIf (gB_BIRA_MC000_flag) Then
            m_PMname = m_PMname ''''No Change, follow the table
        End If
        
        '---------------------------------------------------------------
        '20151120 update
        ''''<MUST> Change the pattern as Index, for correct in Debug Datalog
        m_idx = CpuMbist_PMode_Index(m_PMname)
        m_PM_pattname = CpuMbist_PMode.Category(m_idx).PattName
        m_PM_pattRawname = CpuMbist_PMode.Category(m_idx).PattRawName
        '---------------------------------------------------------------
        
        If (m_1st_elemStr Like UCase("CpuPllInit*")) Then
            If (UCase(m_TestInstName) Like UCase("CpuPllInit*_NRS_*")) Then
                ''''do NOT change PMode
                ''''Here is _NRS case
                For i = 0 To UBound(m_TestInstName_Arr)
                    m_elemStr = m_TestInstName_Arr(i)
                    If (UCase(m_elemStr) Like UCase("MC*")) Then
                        m_elemStrR = Replace(m_elemStr, "MC", "")
                        If (IsNumeric(m_elemStrR) = True) Then
                            m_PMname = m_elemStr + "_NRS"
                        End If
                        ''''<MUST> Change the pattern as Index
                        m_idx = CpuMbist_PMode_Index(m_PMname)
                        m_PM_pattname = CpuMbist_PMode.Category(m_idx).PattName
                        m_PM_pattRawname = CpuMbist_PMode.Category(m_idx).PattRawName
                    End If
                    If (i = 0) Then
                        m_tmpStr = m_elemStr + m_grpnameR ''''Add Group Name
                    ElseIf (i = UBound(m_TestInstName_Arr)) Then
                        m_tmpStr = m_tmpStr + "_" + m_BMname + "_" + m_elemStr ''''Add BMode Name
                    ElseIf (i > 0) Then
                        m_tmpStr = m_tmpStr + "_" + m_elemStr
                    End If
                Next i
                '''m_idx = 0 ''reset, debug
            Else
                ''''Here is Non NRS case
                For i = 0 To UBound(m_TestInstName_Arr)
                    m_elemStr = m_TestInstName_Arr(i)
                    If (UCase(m_elemStr) Like UCase("MC*")) Then
                        m_elemStrR = Replace(m_elemStr, "MC", "")
                        If (IsNumeric(m_elemStrR) = True) Then
                            If (m_PMname Like UCase("*_NRS")) Then
                                ''''NRS case
                                m_elemStr = Replace(m_PMname, "_NRS", "")
                            Else
                                m_elemStr = m_PMname
                            End If
                        End If
                        ''''<MUST> Change the pattern as Index
                        m_idx = CpuMbist_PMode_Index(m_PMname)
                        m_PM_pattname = CpuMbist_PMode.Category(m_idx).PattName
                        m_PM_pattRawname = CpuMbist_PMode.Category(m_idx).PattRawName
                    End If
                    If (i = 0) Then
                        m_tmpStr = m_elemStr + m_grpnameR ''''Add Group,PMode Name
                    ElseIf (i = UBound(m_TestInstName_Arr)) Then
                        m_tmpStr = m_tmpStr + "_" + m_BMname + "_" + m_elemStr ''''Add BMode Name
                    ElseIf (i > 0) Then
                        m_tmpStr = m_tmpStr + "_" + m_elemStr
                    End If
                Next i
            End If

        ElseIf (m_1st_elemStr Like UCase("CpuBlockSel*")) Then
            ''''Block Mode
            For i = 0 To UBound(m_TestInstName_Arr)
                m_elemStr = m_TestInstName_Arr(i)
                If (UCase(m_elemStr) Like UCase("B*")) Then
                    m_elemStrR = Replace(m_elemStr, "B", "")
                    If (IsNumeric(m_elemStrR) = True) Then
                        m_elemStr = m_BMname
                    End If
                End If
                If (i = 0) Then m_tmpStr = m_elemStr + m_grpnameR + "_" + m_PMname ''''Add Group,PMode Name
                If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
            Next i
        Else
            If (UCase(m_TestInstName) Like UCase("*_IN*_*")) Then
                ''''With *_IN02/03/*_*, SmtInit TestInstance, Add Group,PMode_Block Name
                For i = 0 To UBound(m_TestInstName_Arr)
                    m_elemStr = m_TestInstName_Arr(i)
                    If (i = 0) Then m_tmpStr = m_elemStr + m_grpnameR + "_" + m_PMname + "_" + m_BMname  ''''Add Group,PMode_Block Name
                    If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
                Next i
            Else
                ''''PayLoad or Others
                For i = 0 To UBound(m_TestInstName_Arr)
                    m_elemStr = m_TestInstName_Arr(i)
                    If (i = 0) Then m_tmpStr = m_elemStr + "_" + m_PMname + "_" + m_BMname ''''Add PMode_Block Name
                    If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
                Next i
            End If
    
        End If
        
    Else
        ''''-------------------------------------------------------------------------
        '''' Non "BIRA_ALL" Case
        ''''-------------------------------------------------------------------------
        If (m_1st_elemStr Like UCase("CpuPllInit*")) Then
            ''''Performance Mode
            If (UCase(m_grpname) Like UCase("PCM_*")) Then
                ''''Here is the case which *IN00* in inside the Block Loop
                ''''20151110 special case for PCM
                If (UCase(m_TestInstName) Like UCase("*_NRS_*")) Then
                    ''''do NOT change PMode
                    ''''Here is _NRS case
                    For i = 0 To UBound(m_TestInstName_Arr)
                        m_elemStr = m_TestInstName_Arr(i)
                        If (UCase(m_elemStr) Like UCase("MC*")) Then
                            m_elemStrR = Replace(m_elemStr, "MC", "")
                            If (IsNumeric(m_elemStrR) = True) Then
                                m_PMname = m_elemStr + "_NRS"
                            End If
                            m_idx = CpuMbist_PMode_Index(m_PMname)
                            m_PM_pattname = CpuMbist_PMode.Category(m_idx).PattName
                            m_PM_pattRawname = CpuMbist_PMode.Category(m_idx).PattRawName
                        End If
                        If (i = 0) Then
                            m_tmpStr = m_elemStr + m_grpnameR ''''Add Group Name
                        ElseIf (i = UBound(m_TestInstName_Arr)) Then
                            m_tmpStr = m_tmpStr + "_" + m_BMname + "_" + m_elemStr ''''Add BMode Name
                        ElseIf (i > 0) Then
                            m_tmpStr = m_tmpStr + "_" + m_elemStr
                        End If
                    Next i
                    '''m_idx = 0 ''reset, debug
                Else
                    For i = 0 To UBound(m_TestInstName_Arr)
                        m_elemStr = m_TestInstName_Arr(i)
                        If (UCase(m_elemStr) Like UCase("MC*")) Then
                            m_elemStrR = Replace(m_elemStr, "MC", "")
                            If (IsNumeric(m_elemStrR) = True) Then
                                If (m_PMname Like UCase("*_NRS")) Then
                                    ''''NRS case
                                    m_elemStr = Replace(m_PMname, "_NRS", "")
                                Else
                                    m_elemStr = m_PMname
                                End If
                            End If
                        End If
                        If (i = 0) Then
                            m_tmpStr = m_elemStr + m_grpnameR ''''Add Group,PMode Name
                        ElseIf (i = UBound(m_TestInstName_Arr)) Then
                            m_tmpStr = m_tmpStr + "_" + m_BMname + "_" + m_elemStr ''''Add BMode Name
                        ElseIf (i > 0) Then
                            m_tmpStr = m_tmpStr + "_" + m_elemStr
                        End If
                    Next i
                End If
            Else
                ''''default rule
                For i = 0 To UBound(m_TestInstName_Arr)
                    m_elemStr = m_TestInstName_Arr(i)
                    If (UCase(m_elemStr) Like UCase("MC*")) Then
                        m_elemStrR = Replace(m_elemStr, "MC", "")
                        If (IsNumeric(m_elemStrR) = True) Then
                            If (m_PMname Like UCase("*_NRS")) Then
                                ''''NRS case
                                m_elemStr = Replace(m_PMname, "_NRS", "")
                            Else
                                m_elemStr = m_PMname
                            End If
                        End If
                    End If
                    If (i = 0) Then m_tmpStr = m_elemStr + m_grpnameR ''''Add Group,PMode Name
                    If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
                Next i
            End If
        ElseIf (m_1st_elemStr Like UCase("CpuBlockSel*")) Then
            ''''Block Mode
            For i = 0 To UBound(m_TestInstName_Arr)
                m_elemStr = m_TestInstName_Arr(i)
                If (UCase(m_elemStr) Like UCase("B*")) Then
                    m_elemStrR = Replace(m_elemStr, "B", "")
                    If (IsNumeric(m_elemStrR) = True) Then
                        m_elemStr = m_BMname
                    End If
                End If
                If (i = 0) Then m_tmpStr = m_elemStr + m_grpnameR + "_" + m_PMname ''''Add Group,PMode Name
                If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
            Next i
        Else
            If (UCase(m_TestInstName) Like UCase("*_IN*_*")) Then
                ''''With *_IN02/03/*_*, SmtInit TestInstance, Add Group,PMode_Block Name
                For i = 0 To UBound(m_TestInstName_Arr)
                    m_elemStr = m_TestInstName_Arr(i)
                    If (i = 0) Then m_tmpStr = m_elemStr + m_grpnameR + "_" + m_PMname + "_" + m_BMname  ''''Add Group,PMode_Block Name
                    If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
                Next i
            Else
                ''''PayLoad or Others
                For i = 0 To UBound(m_TestInstName_Arr)
                    m_elemStr = m_TestInstName_Arr(i)
                    If (i = 0) Then m_tmpStr = m_elemStr + "_" + m_PMname + "_" + m_BMname ''''Add PMode_Block Name
                    If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
                Next i
            End If
    
        End If

    End If

    If (Mid(m_tmpStr, 1, 1) = "_") Then
        m_tmpStr = Mid(m_tmpStr, 2, Len(m_tmpStr) - 1) ''''remove the first underscore '_'
    End If
    If (Mid(m_tmpStr, Len(m_tmpStr), 1) = "_") Then
        m_tmpStr = Mid(m_tmpStr, 1, Len(m_tmpStr) - 1) ''''remove the last underscore '_'
    End If
    m_testName = m_tmpStr
    ''''-------------------------------------------------------------------------------------------------
    ''''   Set TestName to instance name for debug print CHWu 102615
    ''''-------------------------------------------------------------------------------------------------
    G_TestName = m_tmpStr

    ''''-------------------------------------------------------------------------------------------------
    '''' Set the specificied DC/AC Category Name
    ''''-------------------------------------------------------------------------------------------------
    ''''Get the current TestInstance Context
    Dim m_DCCategory As String
    Dim m_DCSelector As String
    Dim m_ACCategory As String
    Dim m_ACSelector As String
    Dim m_TimeSetSheet As String
    Dim m_EdgeSetSheet As String
    Dim m_LevelsSheet As String
    Dim m_tmpPMname As String

    ''''20151109
    ''''Use the local module private global variable to be flexible if it could be used anywhere in this Module. (Just in case)
    Call TheExec.DataManager.GetInstanceContext(gm_DCCategory, gm_DCSelector, _
                                                gm_ACCategory, gm_ACSelector, _
                                                gm_TimeSetSheet, gm_EdgeSetSheet, _
                                                gm_LevelsSheet, "")

    ''''Here it is transfered to local variable to meet the following statement.
    m_DCCategory = gm_DCCategory
    m_DCSelector = gm_DCSelector
    m_ACCategory = gm_ACCategory
    m_ACSelector = gm_ACSelector
    m_TimeSetSheet = gm_TimeSetSheet
    m_EdgeSetSheet = gm_EdgeSetSheet
    m_LevelsSheet = gm_LevelsSheet

    ''''Initialize
    m_DCCate = m_DCCategory
    m_ACCate = m_ACCategory

    m_tailStr = UCase(Right(m_TestInstName, 5))
    
    ''''-----------------------------------------------------------
    ''''to decide the DC/AC Category and Selector
    ''''-----------------------------------------------------------
    ''''<NOTICE> Here using "m_tmpPMname" to Replace AC/DC Category Name
    m_tmpPMname = m_PMname
    ''''Update PMode 'MCxxx_NRS' to 'MCxxx'
    If (m_PMname Like UCase("*_NRS")) Then ''''NRS case
        m_tmpPMname = Replace(m_PMname, "_NRS", "")
        
        ''''Special case, 20151110 update D64
        If (UCase(m_grpname) Like UCase("*D64*HardDefect*")) Then
            m_tmpPMname = "MrgHardD64"
        ElseIf (UCase(m_grpname) Like UCase("*D64*BitCell*")) Then
            m_tmpPMname = "MrgBitcellD64"
        ElseIf (UCase(m_grpname) Like UCase("*HardDefect*")) Then
            m_tmpPMname = "MrgHard"
        ElseIf (UCase(m_grpname) Like UCase("*BitCell*")) Then
            m_tmpPMname = "MrgBitcell"
        End If
        
        ''''20151111 Update
        If (UCase(m_grpname) = UCase("BIRA_ALL")) Then
            If (gB_BIRA_MC000_NRS_flag) Then
                m_tmpPMname = "MrgHard"
                If (UCase(m_TestInstName) Like UCase("*BitCell*")) Then
                    m_tmpPMname = "MrgBitcell"
                End If
            ''ElseIf (gB_BIRA_MC010_NRS_flag Or gB_BIRA_MC051_NRS_flag) Then
            ElseIf (gB_BIRA_MC010_NRS_flag) Then
                m_tmpPMname = "MrgHardD64"
                If (UCase(m_TestInstName) Like UCase("*BitCell*")) Then
                    m_tmpPMname = "MrgBitcellD64"
                End If
            End If
        End If
    End If

    For i = 0 To UBound(CpuMbist_VoltSet.Category)
        m_VMname = "*_" + UCase(CpuMbist_VoltSet.Category(i).Name)
        If (m_tailStr Like m_VMname) Then
            ''''decide DC/AC Selector
            m_DCSelector = CpuMbist_VoltSet.Category(i).DCSele
            m_ACSelector = CpuMbist_VoltSet.Category(i).ACSele
            
            ''''decide DC Category
            m_tmpStr = CpuMbist_VoltSet.Category(i).DCCate
            
                
            ''''20151111 Update
            If (UCase(m_grpname) = UCase("BIRA_ALL")) Then
                If (gB_BIRA_MC000_NRS_flag) Then
                    m_tmpPMname = "MrgHard"
                    If (UCase(m_TestInstName) Like UCase("*BitCell*")) Then
                        m_tmpPMname = "MrgBitcell"
                    End If
                    m_DCCate = "Cpu" + m_tmpPMname
                ElseIf (gB_BIRA_MC010_NRS_flag) Then
                    m_tmpPMname = "MrgHardD64"
                    If (UCase(m_TestInstName) Like UCase("*BitCell*")) Then
                        m_tmpPMname = "MrgBitcellD64"
                    End If
                    m_DCCate = "Cpu" + m_tmpPMname
                ElseIf (gB_BIRA_MC051_NRS_flag) Then
                    m_tmpPMname = "MC051"
                    m_DCCate = Replace(m_tmpStr, "MC000", m_tmpPMname)
                Else
                    m_tmpPMname = m_PMname
                    m_DCCate = Replace(m_tmpStr, "MC000", m_tmpPMname)
                End If
                
            ElseIf (UCase(m_tmpStr) Like "*MC000*") Then
                m_DCCate = Replace(m_tmpStr, "MC000", m_tmpPMname)
            Else
                ''''Special case, 20151111 update
                If (UCase(m_grpname) Like UCase("*D64*HardDefect*")) Then
                    m_tmpPMname = "MrgHardD64"
                ElseIf (UCase(m_grpname) Like UCase("*D64*BitCell*")) Then
                    m_tmpPMname = "MrgBitcellD64"
                ElseIf (UCase(m_grpname) Like UCase("*HardDefect*")) Then
                    m_tmpPMname = "MrgHard"
                ElseIf (UCase(m_grpname) Like UCase("*BitCell*")) Then
                    m_tmpPMname = "MrgBitcell"
                End If
                
                m_DCCate = "Cpu" + m_tmpPMname
            End If
            
            ''''decide AC Category
            m_tmpStr = CpuMbist_VoltSet.Category(i).ACCate
            If (UCase(m_tmpStr) Like "*MC000*") Then
                m_ACCate = Replace(m_tmpStr, "MC000", m_tmpPMname)
            Else
                ''''Special case, 20151111 update
                If (UCase(m_grpname) Like UCase("*D64*HardDefect*")) Then
                    m_tmpPMname = "MrgHardD64"
                ElseIf (UCase(m_grpname) Like UCase("*D64*BitCell*")) Then
                    m_tmpPMname = "MrgBitcellD64"
                ElseIf (UCase(m_grpname) Like UCase("*HardDefect*")) Then
                    m_tmpPMname = "MrgHard"
                ElseIf (UCase(m_grpname) Like UCase("*BitCell*")) Then
                    m_tmpPMname = "MrgBitcell"
                End If
                m_ACCate = "Cpu" + m_tmpPMname
            End If
            
            Exit For
        End If
    Next i

    ''''<Importance>
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, _
                                     PinLevelsSheet:=m_LevelsSheet, DCCategory:=m_DCCate, DCSelector:=m_DCSelector, _
                                     TimeSetSheet:=m_TimeSetSheet, ACCategory:=m_ACCate, ACSelector:=m_ACSelector, EdgeSetSheet:=m_EdgeSetSheet

    TheHdw.Wait 0.0001 ''''100uS

    ''''<Error Here>
    ''''{-1} : error DMC:0023 : Cannot specify category or selector that is not used in job
    ''''theHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, DCCategory:=m_DCCate, ACCategory:=m_ACCate
    ''''<Workaround>
    ''''Create a dummy Test function auto_Dummy_Category() and its related Dummy Test/TestInstance
    ''''-------------------------------------------------------------------------------------------------
    
    
    ''''-------------------------------------------------------------------------------------------------
    '''' Set the specified BinFlagName
    ''''-------------------------------------------------------------------------------------------------
    m_ChangeFlag = UCase(CpuMbist.Category(m_grpIndex).ChangeFlag) ''''<MUST> use UCase
    If (UCase(m_TestInstName) Like UCase("*_IN0*_*")) Then
        ''''20151109, all INIT patterns should have its default BinFlagName
        ''''do NOT change BinFlagName
        m_tmpStr = ""
        m_ChangeFlag = "F" ''''<MUST>
    ElseIf (m_1st_elemStr Like UCase("CpuBist*")) Then
        m_tmpStr = LCase("CpuBist") ''''<MUST> use LCase
    ElseIf (m_1st_elemStr Like UCase("CpuBira*")) Then
        m_tmpStr = LCase("CpuBira") ''''<MUST> use LCase
    ElseIf (m_1st_elemStr Like UCase("CpuRetMargin*")) Then
        m_tmpStr = LCase("CpuRetMargin") ''''<MUST> use LCase
    Else
        ''''do NOT change BinFlagName
        m_tmpStr = ""
        m_ChangeFlag = "F" ''''<MUST>
    End If
    m_BinFlagName = auto_Mbist_GetBinFlagName(m_ChangeFlag, m_TestInstName_Arr, m_tmpStr, m_last_elemStr, m_PMname, m_BMname, m_testName)
    ''''-------------------------------------------------------------------------------------------------
    
    ''''-------------------------------------------------------------------------------------------------
    ''''Decide which pattern to be Executed
    ''''-------------------------------------------------------------------------------------------------
    Dim m_status As Boolean
    Dim m_patArr() As String
    Dim m_patcount As Long
    If (m_1st_elemStr Like UCase("CpuPllInit*")) Then
        ''''Performance Mode
        m_pattname = m_PM_pattname
        m_pattRawname = m_PM_pattRawname
    ElseIf (m_1st_elemStr Like UCase("CpuBlockSel*")) Then
        ''''Block Mode
        m_pattname = m_BM_pattname
        m_pattRawname = m_BM_pattRawname
    Else
        ''''20151105 update
        ''''Other Cases do NOT update pattern (check)
        m_pattname = m_pattname
        If (m_pattname <> "") Then
            ''''------------------------------------------------------------------------
            m_status = GetPatListFromPatternSet(m_pattname, m_patArr, m_patcount)
            If (m_status) Then
                If (m_patcount = 1) Then
                    m_pattRawname = m_patArr(0)
                Else
                    ''''there are multi-patterns in the PattSet, it needs another process later on
                    m_pattRawname = m_pattname
                End If
            Else
                ''''there are multi-patterns in the PattSet, it needs another process later on
                m_pattRawname = m_pattname
            End If
            ''''------------------------------------------------------------------------
        Else
            ''''Null Pattern input
        End If
    End If
    ''''-------------------------------------------------------------------------------------------------

    ''''-------------------------------------------------------------------------------------------------
    ''''Summarize All Information Before run the Funtional Test
    ''''-------------------------------------------------------------------------------------------------
    If (showPrint) Then
        PrintDataLog ""
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog Space(25) + funcName
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog "       GroupName:: " + m_grpname + " (CpuMbist.Category Index = " + CStr(m_grpIndex) + ")"
        PrintDataLog "  Loop Parameter:: " + _
                     "LP_DM=" + FormatNumeric(m_LoopIndex_DM, -3) + " , " + _
                     "LP_PM=" + FormatNumeric(m_LoopIndex_PM, -3) + " , " + _
                     "LP_BM=" + FormatNumeric(m_LoopIndex_BM, -3) + " (Current LoopIndex)"
        PrintDataLog "Performance Mode:: " + FormatNumeric(m_PMname, 9) + " , Pattern=" + m_PM_pattname
        PrintDataLog "      Block Mode:: " + FormatNumeric(m_BMname, 9) + " , Pattern=" + m_BM_pattname
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog "     Level_Sheet:: " + m_LevelsSheet
        PrintDataLog "     DC_Category:: " + m_DCCate + ", DC_Selector:: " + m_DCSelector
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog "   TimeSet_Sheet:: " + m_TimeSetSheet
        PrintDataLog "     AC_Category:: " + m_ACCate + ", AC_Selector:: " + m_ACSelector
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog "       Test_Name:: " + m_testName
        PrintDataLog " Performace Patt:: " + m_PM_pattRawname
        PrintDataLog "      Block Patt:: " + m_BM_pattRawname
        If (m_pattname <> "") Then
            PrintDataLog "Executed Pattern:: " + m_pattname + ", RawPatt=" + m_pattRawname
        End If
        PrintDataLog "    Bin_FlagName:: " + m_BinFlagName
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog ""
    End If
    ''''-------------------------------------------------------------------------------------------------

    auto_FuncTest_CpuMbist_Body = 1
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_FuncTest_Mbist_Execute(m_pattname As String, _
                                            m_testName As String, _
                                            m_BinFlagName As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_FuncTest_Mbist_Execute"

    ''''-------------------------------------------------------------------------------------------------
    ''''20151020 Update (Check for Mbist Function)
    ''''-------------------------------------------------------------------------------------------------
    Dim m_tn As Long
    Dim Site As Variant
    
    ''''-------------------------------------------------------------------------------------------------
    ''''20151102, Reset F_Payload Every time before runing payload
    Dim m_flagname As String
    m_flagname = "F_Payload"
    For Each Site In TheExec.Sites.Existing
        TheExec.Sites.Item(Site).FlagState(m_flagname) = logicFalse ''''mean Pass
    Next Site
    gS_currPayload_pattSetName = m_pattname ''''for SONE datalog
    ''''-------------------------------------------------------------------------------------------------
    
    TheHdw.Patterns(m_pattname).start
    TheHdw.Digital.Patgen.HaltWait
    For Each Site In TheExec.Sites
        m_tn = TheExec.Sites.Item(Site).TestNumber
        If TheHdw.Digital.Patgen.PatternBurstPassed = True Then
            Call TheExec.Datalog.WriteFunctionalResult(Site, m_tn, logTestPass, , m_testName)
            If (UCase(m_BinFlagName) <> UCase("Default")) Then
                If (TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicTrue) Then
                    ''''<Important>
                    ''''Because it was Failed on previous test, so it will NOT do any change here.
                Else
                    TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicFalse ''''mean Pass
                End If
            End If
        Else
            ''''Fail/Alarm Case
            Call TheExec.Datalog.WriteFunctionalResult(Site, m_tn, logTestFail, , m_testName)
            TheExec.Sites.Item(Site).TestResult = siteFail ''''20151112 update
            If (UCase(m_BinFlagName) <> UCase("Default")) Then
                TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicTrue ''''mean Fail
            End If
            ''''20151102, for SONE
            TheExec.Sites.Item(Site).FlagState(m_flagname) = logicTrue ''''mean Fail
        End If
    Next Site
    TheExec.Flow.IncrementTestNumber
    ''''-------------------------------------------------------------------------------------------------
    
    Dim m_instName As String
    m_instName = TheExec.DataManager.InstanceName
    If UCase(m_instName) Like UCase("*RING*") Then
        Dim MeasF_Pin As New PinList
        MeasF_Pin.Value = "RINGS_RO_CLK_OUT"
        'Call HardIP_FrequencyMeasure(MeasureF_Pin_SingleEnd, False, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, Pat, Flag_SingleLimit, d_MeasF_Interval, MeasF_WaitTime, MeasF_EventSource)
         Call HardIP_FrequencyMeasure(MeasF_Pin, False, "FFF", 0, 0, 0, m_pattname, True, 0.01)
    End If
    
    auto_FuncTest_Mbist_Execute = 1
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151105 Update
Public Function auto_FuncTest_CpuMbist_CharZ_Body(ByRef pattSet As Pattern, _
                                                  ByRef retINITpattArr() As String, _
                                                  ByRef retPLLPpattArr() As String, _
                                                  ByRef m_testName As String, _
                                                  ByRef m_pattseq As String, _
                                                  ByRef m_pattseqcnt As Long, _
                                                  Optional showPrint As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_FuncTest_CpuMbist_CharZ_Body"

    ''''CpuMbist---------------------------------
    Dim i As Long
    Dim m_tmpStr As String
    Dim m_grpIndex As Long
    Dim m_grpname As String
    Dim m_grpnameR As String
    Dim m_LoopIndex_DM As Long
    Dim m_LoopIndex_PM As Long
    Dim m_LoopIndex_BM As Long
    
    Dim m_PMname As String
    Dim m_PM_pattname As String
    Dim m_PM_pattRawname As String
    Dim m_BMname As String
    Dim m_BM_pattname As String
    Dim m_BM_pattRawname As String
    Dim m_DMname As String
    Dim m_pattname As String
    Dim m_pattRawname As String
    ''''----------------------------------------
    Dim m_TestInstName As String
    Dim m_TestInstName_Arr() As String
    Dim m_1st_elemStr As String
    Dim m_last_elemStr As String
    Dim m_elemStr As String
    Dim m_elemStrR As String
    ''''----------------------------------------
    Dim m_tailStr As String
    Dim m_VMname As String
    Dim m_DCCate As String
    Dim m_ACCate As String
    ''''----------------------------------------
    ''Dim m_BinFlagName As String
    Dim m_ChangeFlag As String
    ''''----------------------------------------
    
    ''''-------------------------------------------------------------------------------------------------
    '''' Get the related parameters of Mbist
    ''''-------------------------------------------------------------------------------------------------
    m_grpname = gS_Mbist_GroupName
    m_grpIndex = gL_Mbist_GroupName_Index
    Call auto_Mbist_GetLoopIndex_DM_PM_BM(m_LoopIndex_DM, m_LoopIndex_PM, m_LoopIndex_BM, False)

    ''''--------------------------------------------------------------
    ''''<NOTICE>
    ''''--------------------------------------------------------------
    ''''LP_DM For-Loop MUST be from 0 in the Flow and <= LCount_DM
    ''''LP_PM For-Loop MUST be from 0 in the Flow and <  LCount_PM
    ''''LP_BM For-Loop MUST be from 0 in the Flow and <  LCount_BM
    ''''--------------------------------------------------------------
    
    m_DMname = UCase(CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).Name)
    If (m_DMname = "NA") Then
        ''''Case: DebugMode is Disable
        m_PMname = UCase(CpuMbist.Category(m_grpIndex).PMode(m_LoopIndex_PM).Name)
        m_PM_pattname = CpuMbist.Category(m_grpIndex).PMode(m_LoopIndex_PM).PattName
        m_PM_pattRawname = CpuMbist.Category(m_grpIndex).PMode(m_LoopIndex_PM).PattRawName
        
        m_BMname = UCase(CpuMbist.Category(m_grpIndex).Block(m_LoopIndex_BM).Name)
        m_BM_pattname = CpuMbist.Category(m_grpIndex).Block(m_LoopIndex_BM).PattName
        m_BM_pattRawname = CpuMbist.Category(m_grpIndex).Block(m_LoopIndex_BM).PattRawName
    Else
        ''''Case: DebugMode is Enable
        ''''<NOTICE> Only use 'm_LoopIndex_DM' to decide the PMode and Block mode
        m_PMname = UCase(CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).PMode)
        m_PM_pattname = CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).PMode_PattName
        m_PM_pattRawname = CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).PMode_PattRawName
        
        m_BMname = UCase(CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).Block)
        m_BM_pattname = CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).Block_PattName
        m_BM_pattRawname = CpuMbist.Category(m_grpIndex).DebugMode(m_LoopIndex_DM).Block_PattRawName
    End If
    ''''-------------------------------------------------------------------------------------------------
    
    ''''-------------------------------------------------------------------------------------------------
    '''' Set the specified TestName
    ''''-------------------------------------------------------------------------------------------------
    m_grpnameR = Replace(m_grpname, "_", "")
    m_TestInstName = Trim(TheExec.DataManager.InstanceName)
    m_TestInstName_Arr = Split(m_TestInstName, "_")
    ReDim Preserve m_TestInstName_Arr(UBound(m_TestInstName_Arr))
    m_1st_elemStr = UCase(m_TestInstName_Arr(0))
    m_last_elemStr = m_TestInstName_Arr(UBound(m_TestInstName_Arr))
    m_testName = ""
    m_tmpStr = ""
    If (m_1st_elemStr Like UCase("CpuPllInit*")) Then
        ''''Performance Mode
        For i = 0 To UBound(m_TestInstName_Arr)
            m_elemStr = m_TestInstName_Arr(i)
            If (UCase(m_elemStr) Like UCase("MC*")) Then
                m_elemStrR = Replace(m_elemStr, "MC", "")
                If (IsNumeric(m_elemStrR) = True) Then
                    If (m_PMname Like UCase("*_NRS")) Then
                        ''''NRS case
                        m_elemStr = Replace(m_PMname, "_NRS", "")
                    Else
                        m_elemStr = m_PMname
                    End If
                End If
            End If
            If (i = 0) Then m_tmpStr = m_elemStr + m_grpnameR ''''Add Group,PMode Name
            If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
        Next i
    ElseIf (m_1st_elemStr Like UCase("CpuBlockSel*")) Then
        ''''Block Mode
        For i = 0 To UBound(m_TestInstName_Arr)
            m_elemStr = m_TestInstName_Arr(i)
            If (UCase(m_elemStr) Like UCase("B*")) Then
                m_elemStrR = Replace(m_elemStr, "B", "")
                If (IsNumeric(m_elemStrR) = True) Then
                    m_elemStr = m_BMname
                End If
            End If
            If (i = 0) Then m_tmpStr = m_elemStr + m_grpnameR + "_" + m_PMname ''''Add Group,PMode Name
            If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
        Next i
    Else
        If (UCase(m_TestInstName) Like UCase("*_IN*_*")) Then
            ''''With *_IN*_*, init TestInstance, Add Group,PMode_Block Name
            For i = 0 To UBound(m_TestInstName_Arr)
                m_elemStr = m_TestInstName_Arr(i)
                If (i = 0) Then m_tmpStr = m_elemStr + m_grpnameR + "_" + m_PMname + "_" + m_BMname  ''''Add Group,PMode_Block Name
                If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
            Next i
        Else
            ''''PayLoad or Others
            For i = 0 To UBound(m_TestInstName_Arr)
                m_elemStr = m_TestInstName_Arr(i)
                If (i = 0) Then m_tmpStr = m_elemStr + "_" + m_PMname + "_" + m_BMname ''''Add PMode_Block Name
                If (i > 0) Then m_tmpStr = m_tmpStr + "_" + m_elemStr
            Next i
        End If

    End If
    
    If (Mid(m_tmpStr, 1, 1) = "_") Then
        m_tmpStr = Mid(m_tmpStr, 2, Len(m_tmpStr) - 1) ''''remove the first underscore '_'
    End If
    If (Mid(m_tmpStr, Len(m_tmpStr), 1) = "_") Then
        m_tmpStr = Mid(m_tmpStr, 1, Len(m_tmpStr) - 1) ''''remove the last underscore '_'
    End If
    m_testName = m_tmpStr
    
    ''''-------------------------------------------------------------------------------------------------
    ''''   Set TestName to instance name for debug print CHWu 102615
    ''''-------------------------------------------------------------------------------------------------
    G_TestName = m_tmpStr
    
    ''''-------------------------------------------------------------------------------------------------
    '''' Set the specificied DC/AC Category Name
    ''''-------------------------------------------------------------------------------------------------
    ''''Get the current TestInstance Context
    Dim m_DCCategory As String
    Dim m_DCSelector As String
    Dim m_ACCategory As String
    Dim m_ACSelector As String
    Dim m_TimeSetSheet As String
    Dim m_EdgeSetSheet As String
    Dim m_LevelsSheet As String
    Dim m_tmpPMname As String

    ''''20151109
    ''''Use the local module private global variable to be flexible if it could be used anywhere in this Module. (Just in case)
    Call TheExec.DataManager.GetInstanceContext(gm_DCCategory, gm_DCSelector, _
                                                gm_ACCategory, gm_ACSelector, _
                                                gm_TimeSetSheet, gm_EdgeSetSheet, _
                                                gm_LevelsSheet, "")

    ''''Here it is transfered to local variable to meet the following statement.
    m_DCCategory = gm_DCCategory
    m_DCSelector = gm_DCSelector
    m_ACCategory = gm_ACCategory
    m_ACSelector = gm_ACSelector
    m_TimeSetSheet = gm_TimeSetSheet
    m_EdgeSetSheet = gm_EdgeSetSheet
    m_LevelsSheet = gm_LevelsSheet

    ''''Initialize
    m_DCCate = m_DCCategory
    m_ACCate = m_ACCategory

    m_tailStr = UCase(Right(m_TestInstName, 5))
    
    ''''-----------------------------------------------------------
    ''''to decide the DC/AC Category and Selector
    ''''-----------------------------------------------------------
    ''''<NOTICE> Here using "m_tmpPMname" to Replace AC/DC Category Name
    m_tmpPMname = m_PMname
    ''''Update PMode 'MCxxx_NRS' to 'MCxxx'
    If (m_PMname Like UCase("*_NRS")) Then ''''NRS case
        m_tmpPMname = Replace(m_PMname, "_NRS", "")
        
        ''''Special case, 20151110 update D64
        If (UCase(m_grpname) Like UCase("*HardDefectD64*")) Then
            m_tmpPMname = "MrgHardD64"
        ElseIf (UCase(m_grpname) Like UCase("*BitCellD64*")) Then
            m_tmpPMname = "MrgBitcellD64"
        ElseIf (UCase(m_grpname) Like UCase("*HardDefect*")) Then
            m_tmpPMname = "MrgHard"
        ElseIf (UCase(m_grpname) Like UCase("*BitCell*")) Then
            m_tmpPMname = "MrgBitcell"
        End If
    End If

    For i = 0 To UBound(CpuMbist_VoltSet.Category)
        m_VMname = "*_" + UCase(CpuMbist_VoltSet.Category(i).Name)
        If (m_tailStr Like m_VMname) Then
            ''''decide DC/AC Selector
            m_DCSelector = CpuMbist_VoltSet.Category(i).DCSele
            m_ACSelector = CpuMbist_VoltSet.Category(i).ACSele
            
            ''''decide DC Category
            m_tmpStr = CpuMbist_VoltSet.Category(i).DCCate
            If (UCase(m_tmpStr) Like "*MC000*") Then
                m_DCCate = Replace(m_tmpStr, "MC000", m_tmpPMname)
            Else
                m_DCCate = m_tmpStr
            End If
            
            ''''decide AC Category
            m_tmpStr = CpuMbist_VoltSet.Category(i).ACCate
            If (UCase(m_tmpStr) Like "*MC000*") Then
                m_ACCate = Replace(m_tmpStr, "MC000", m_tmpPMname)
            Else
                m_ACCate = m_tmpStr
            End If
            
            Exit For
        End If
    Next i

    ''''<Importance>
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, _
                                     PinLevelsSheet:=m_LevelsSheet, DCCategory:=m_DCCate, DCSelector:=m_DCSelector, _
                                     TimeSetSheet:=m_TimeSetSheet, ACCategory:=m_ACCate, ACSelector:=m_ACSelector, EdgeSetSheet:=m_EdgeSetSheet

    TheHdw.Wait 0.0001 ''''100uS
    
    ''''<Error Here>
    ''''{-1} : error DMC:0023 : Cannot specify category or selector that is not used in job
    ''''theHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, DCCategory:=m_DCCate, ACCategory:=m_ACCate
    ''''<Workaround>
    ''''Create a dummy Test function auto_Dummy_Category() and its related Dummy Test/TestInstance
    ''''-------------------------------------------------------------------------------------------------
    
    ''''No need the BinFlagName in CharZ flow
''''    ''''-------------------------------------------------------------------------------------------------
''''    '''' Set the specified BinFlagName
''''    ''''-------------------------------------------------------------------------------------------------
''''    m_ChangeFlag = UCase(CpuMbist.Category(m_grpIndex).ChangeFlag) ''''<MUST> use UCase
''''    If (m_1st_elemStr Like UCase("CpuBist*")) Then
''''        m_tmpStr = LCase("CpuBist") ''''<MUST> use LCase
''''    ElseIf (m_1st_elemStr Like UCase("CpuBira*")) Then
''''        m_tmpStr = LCase("CpuBira") ''''<MUST> use LCase
''''    ElseIf (m_1st_elemStr Like UCase("CpuRetMargin*")) Then
''''        m_tmpStr = LCase("CpuRetMargin") ''''<MUST> use LCase
''''    Else
''''        ''''do NOT change BinFlagName
''''        m_tmpStr = ""
''''        m_ChangeFlag = "F" ''''<MUST>
''''    End If
''''    m_BinFlagName = auto_Mbist_GetBinFlagName(m_ChangeFlag, m_TestInstName_Arr, m_tmpStr, m_last_elemStr, m_PMname, m_BMname, m_TestName)
''''    ''''-------------------------------------------------------------------------------------------------
    
    ''''-------------------------------------------------------------------------------------------------
    ''''Decide which pattern to be Executed
    ''''Decomposite the input pattSet
    ''''-------------------------------------------------------------------------------------------------
    Dim m_patcnt As Long         ''''Number of patterns in set
    Dim m_patRawname() As String ''''Raw pattern name data
    Dim m_patt As Variant
    Dim m_pattU As String
    Dim m_initpatcnt As Long
    Dim m_pllppatcnt As Long
    
    ''''Only decomposite 1st layer
    m_patRawname = TheExec.DataManager.Raw.GetPatternsInSet(pattSet, m_patcnt)

    ReDim retINITpattArr(20)
    ReDim retPLLPpattArr(20)
    m_initpatcnt = 0
    m_pllppatcnt = 0
    For Each m_patt In m_patRawname
        m_pattU = UCase(m_patt)
        
        If (m_pattU Like UCase("*_IN00_*_MC*_NRS*")) Then ''''unchange
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1
        
        ElseIf (m_pattU Like UCase("*_IN00_*_MC*")) Then
            retINITpattArr(m_initpatcnt) = m_PM_pattname ''''use the updated MCxxx pattern (PM)
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN01_*_B*")) Then
            retINITpattArr(m_initpatcnt) = m_BM_pattname ''''use the updated B00...B31 pattern (BM)
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN02_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN03_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN04_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1
            
        ElseIf (m_pattU Like UCase("*_IN05_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN06_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN07_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN08_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN09_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1

        ElseIf (m_pattU Like UCase("*_IN10_*")) Then
            retINITpattArr(m_initpatcnt) = m_patt
            m_initpatcnt = m_initpatcnt + 1
            
        ElseIf (m_pattU Like UCase("*_PLLP_*")) Then ''''unchange for PayLoad pattern
            retPLLPpattArr(m_pllppatcnt) = m_patt
            m_pllppatcnt = m_pllppatcnt + 1
            
        End If
    Next m_patt
    
    ''''<MUST> Set the Final Dimension
    ReDim Preserve retINITpattArr(m_initpatcnt - 1)
    ReDim Preserve retPLLPpattArr(m_pllppatcnt - 1)
    ''''-------------------------------------------------------------------------------------------------
    
    ''''20151107 New
    ''''-------------------------------------------------------------------------------------------------
    ''''Get the pattern sequence of the patternSet to (INIT0,INIT1,...,PLLP1,...)
    ''''-------------------------------------------------------------------------------------------------
    m_pattseq = ""
    m_pattseqcnt = 0
    For i = 0 To UBound(retINITpattArr)
        m_pattseqcnt = m_pattseqcnt + 1
        If (i = 0) Then m_pattseq = retINITpattArr(0)
        If (i > 0) Then
            m_pattseq = m_pattseq + "," + retINITpattArr(i)
        End If
    Next i
    
    For i = 0 To UBound(retPLLPpattArr)
        m_pattseqcnt = m_pattseqcnt + 1
        If (m_pattseq = "") Then
            m_pattseq = retPLLPpattArr(0)
        Else
            m_pattseq = m_pattseq + "," + retPLLPpattArr(i)
        End If
    Next i
    
    ''''-------------------------------------------------------------------------------------------------
    ''''Summarize All Information Before run the Funtional Test
    ''''-------------------------------------------------------------------------------------------------
    If (showPrint) Then
        PrintDataLog ""
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog Space(25) + funcName
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog "       GroupName:: " + m_grpname + " (CpuMbist.Category Index = " + CStr(m_grpIndex) + ")"
        PrintDataLog "  Loop Parameter:: " + _
                     "LP_DM=" + FormatNumeric(m_LoopIndex_DM, -3) + " , " + _
                     "LP_PM=" + FormatNumeric(m_LoopIndex_PM, -3) + " , " + _
                     "LP_BM=" + FormatNumeric(m_LoopIndex_BM, -3) + " (Current LoopIndex)"
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog "     Level_Sheet:: " + m_LevelsSheet
        PrintDataLog "     DC_Category:: " + m_DCCate + ", DC_Selector:: " + m_DCSelector
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog "   TimeSet_Sheet:: " + m_TimeSetSheet
        PrintDataLog "     AC_Category:: " + m_ACCate + ", AC_Selector:: " + m_ACSelector
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog "       Test_Name:: " + m_testName
        PrintDataLog "--------------------------------------------------------------------------------------"
        For i = 0 To UBound(retINITpattArr)
            m_tmpStr = FormatNumeric("INIT Pattern_" + Format(i, "00") + ":: ", 19) + retINITpattArr(i)
            PrintDataLog m_tmpStr
        Next i
        PrintDataLog "--------------------------------------------------------------------------------------"
        For i = 0 To UBound(retPLLPpattArr)
            m_tmpStr = FormatNumeric("PLLP Pattern_" + Format(i, "00") + ":: ", 19) + retPLLPpattArr(i)
            PrintDataLog m_tmpStr
        Next i
        PrintDataLog "--------------------------------------------------------------------------------------"
        PrintDataLog ""
    End If
    ''''-------------------------------------------------------------------------------------------------

    auto_FuncTest_CpuMbist_CharZ_Body = 1
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_FuncTest_Mbist_PCM_LoopV_Execute(m_pattname As String, _
                                                      m_freqPattSet As String, _
                                                      m_testName As String, _
                                                      Optional m_BinFlagName As String = "") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_FuncTest_Mbist_PCM_LoopV_Execute"

    ''''-------------------------------------------------------------------------------------------------
    ''''20151110 Update (Check for Mbist PCM FuncTest and FreqMeas by Loop Voltage)
    ''''-------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim m_tn As Long
    Dim Site As Variant

    ''''For Multiple Voltage Search
    Dim m_pwrpin As String
    Dim m_vstart As Double
    Dim m_vstop As Double
    Dim m_vstep As Double
    Dim m_enable As String
    Dim m_value As Double
    Dim m_tmpName As String
    Dim m_tmpNameFreq As String
    Dim m_valStr As String
    Dim m_tmpStr As String

    Dim k As Long
    Dim m_fpatname As String
    Dim m_elemStr As String
    Dim m_elemStrR As String
    Dim m_elemStrArr() As String
    Dim m_freqRingStr As String

    m_tmpName = m_testName
    m_tmpNameFreq = m_testName

    Dim MeasF_Pin As New PinList
    MeasF_Pin.Value = "RINGS_RO_CLK_OUT"

    Dim m_fstatus As Boolean
    Dim m_fpatcnt As Long
    Dim m_freqPattArr() As String

    If (m_freqPattSet <> "") Then
        TheHdw.Patterns(m_freqPattSet).Load ''''<MUST>
        ''''--------------------------------------------------
        ''''Process Freq Pattern Set
        ''''--------------------------------------------------
        m_fstatus = GetPatListFromPatternSet(m_freqPattSet, m_freqPattArr, m_fpatcnt)
        ''''ReDim Preserve m_freqPattArr(m_fpatcnt - 1) ''''check
        ''''20151111 Add
        If (True) Then
            PrintDataLog "FreqMeas PinName:: " + MeasF_Pin.Value
            PrintDataLog "FreqMeas Pattern:: " + m_freqPattSet
            For i = 0 To UBound(m_freqPattArr)
                PrintDataLog Space(19) + "RawPatt(" + Format(i, "00") + ") = " + m_freqPattArr(i)
            Next i
        End If
    End If

    For i = 0 To UBound(CpuMbist_Power.Category)
        m_pwrpin = CpuMbist_Power.Category(i).pwrpin
        m_vstart = CpuMbist_Power.Category(i).V_Start
        m_vstop = CpuMbist_Power.Category(i).V_Stop
        m_vstep = CpuMbist_Power.Category(i).V_Step
        m_enable = UCase(CpuMbist_Power.Category(i).Enable)
        
        If (m_vstart > m_vstop) Then
            m_vstep = -1# * Abs(m_vstep)
        Else
            m_vstep = Abs(m_vstep)
        End If
        
        If (m_enable = "Y") Then
            m_vstop = m_vstop + 0.001 * m_vstep ''''<MUST>
            For m_value = m_vstart To m_vstop Step m_vstep
                ''''Update testName by force Voltage
                m_valStr = Format(m_value * 1000, "0###") + "mV" ''''was "####'
                m_testName = m_tmpName + "_" + m_valStr

                ''''setup power value
                TheHdw.DCVS.Pins(m_pwrpin).Voltage.Main.Value = m_value
                TheHdw.Wait 0.00001 ''''10us

                If (m_pattname <> "") Then
                    Call auto_run_NewMbist_FuncStart_test(m_pattname, m_testName, "", True) ''''No need set BinFlagName
                End If

                m_elemStr = ""
                m_freqRingStr = ""
                If (m_freqPattSet <> "") Then
                    For j = 0 To UBound(m_freqPattArr)
                        m_fpatname = m_freqPattArr(j)
                        m_elemStrArr = Split(m_fpatname, "_")
                        
                        If (UCase(m_fpatname) Like UCase("*_SI_R*")) Then
                            For k = 0 To UBound(m_elemStrArr)
                                m_elemStr = UCase(m_elemStrArr(k))
                                If (m_elemStr Like UCase("R*")) Then '''' ex 'R00'....'R15'
                                    m_elemStrR = Replace(m_elemStr, "R", "")
                                    If (IsNumeric(m_elemStrR)) Then
                                        m_freqRingStr = m_elemStr
                                        Exit For
                                    End If
                                End If
                            Next k
                        Else
                            ''''do Nothing
                            m_freqRingStr = "R" + Format(j, "00")
                        End If
                        m_testName = m_tmpNameFreq + "_" + m_freqRingStr + "_" + m_valStr
                        Call auto_run_NewMbist_FuncStart_test(m_fpatname, m_testName, "", False) ''''No need set BinFlagName
                        
                        m_testName = m_tmpNameFreq + "_FRQ_" + m_freqRingStr + "_" + m_valStr
                        ''Call HardIP_FrequencyMeasure(MeasureF_Pin_SingleEnd, False, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, Pat, Flag_SingleLimit, d_MeasF_Interval, MeasF_WaitTime, MeasF_EventSource)
                        Call HardIP_FrequencyMeasure_NewMbist(MeasF_Pin, False, "FFF", 0, 0, CInt(j), m_freqPattArr(j), True, 0.01, 0.0001, , m_testName, True, True)
                    Next j
                End If
                TheExec.Datalog.WriteComment ""
            Next m_value
        End If ''''end of If (m_enable = "Y") Then
    Next i
    
    auto_FuncTest_Mbist_PCM_LoopV_Execute = 1
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151111 New for NewMbist Structure
Public Function auto_run_NewMbist_FuncStart_test(m_pattname As String, m_testName As String, Optional m_BinFlagName As String = "", Optional m_logFuncPass As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_run_NewMbist_FuncStart_test"
    
    Dim Site As Variant
    Dim m_tn As Long
    
    TheHdw.Patterns(m_pattname).start
    TheHdw.Digital.Patgen.HaltWait
    
    ''''-------------------------------------------------------------------------------------------------
    For Each Site In TheExec.Sites
        m_tn = TheExec.Sites.Item(Site).TestNumber
        If TheHdw.Digital.Patgen.PatternBurstPassed = True Then
            If (m_logFuncPass) Then Call TheExec.Datalog.WriteFunctionalResult(Site, m_tn, logTestPass, , m_testName)
            If (m_BinFlagName <> "") Then
                If (UCase(m_BinFlagName) <> UCase("Default")) Then
                    If (TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicTrue) Then
                        ''''<Important>
                        ''''Because it was Failed on previous test, so it will NOT do any change here.
                    Else
                        TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicFalse ''''mean Pass
                    End If
                End If
            End If
        Else
            ''''Fail/Alarm Case
            Call TheExec.Datalog.WriteFunctionalResult(Site, m_tn, logTestFail, , m_testName)
            TheExec.Sites.Item(Site).TestResult = siteFail ''''20151112 update
            If (m_BinFlagName <> "") Then
                If (UCase(m_BinFlagName) <> UCase("Default")) Then
                    TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicTrue ''''mean Fail
                End If
            End If
        End If
    Next Site
    TheExec.Flow.IncrementTestNumber
    ''''-------------------------------------------------------------------------------------------------

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''It's copy from LIB_HardIP and update as need here
Public Function HardIP_FrequencyMeasure_NewMbist(FreqMeasPins As PinList, IsDifferentialPin As Boolean, TestLimitByPin_VFI As String, Freq_LowLimit As Double, Freq_HighLimit As Double, TestSeqNum As Integer, Pat As Variant, Flag_SingleLimit As Boolean, _
        d_FreqMeasInterval As Double, Optional MeasFreqWaitTime As String = "", Optional MeasF_EventSource As FreqCtrEventSrcSel, Optional m_testName As String = "", Optional LoLmt_NA_Flag As Boolean = False, Optional HiLmt_NA_Flag As Boolean = False)
        
On Error GoTo errHandler
    Dim funcName As String:: funcName = "HardIP_FrequencyMeasure_NewMbist"

    Dim p As Long
    Dim MeasFreq As New PinListData
    
    Call Freq_MeasFreqSetup(FreqMeasPins, d_FreqMeasInterval, MeasF_EventSource)   '' 20150621 - default d_FreqMeasInterval = 0.01
    '' 20150623 - Add Customize Wait Time
    Call HardIP_Freq_MeasFreqStart(FreqMeasPins, d_FreqMeasInterval, MeasFreq, MeasFreqWaitTime)       '' 20150621 - default d_FreqMeasInterval = 0.01

    Dim TestNameInput As String
    TestNameInput = "Freq_meas_"
    
    If m_testName <> "" Then
       TestNameInput = m_testName
    End If

    If Flag_SingleLimit = True Then
        If LCase(TheExec.CurrentJob) Like "*char*" Then
            TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , ForceResults:=tlForceFlow
        Else
            '' 20151102 - Modify test name as Chihome
            '' TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput + CStr(TestSeqNum) + "_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), ForceResults:=tlForceNone
            If UCase(TestNameInput) Like UCase("*CPU*") Then
                ''TestNameInput = TestNameInput + "_FRQ" + Format(TestSeqNum, "00") ''''has been process outside
                Call UpdateDLogColumns(Len(TestNameInput) + 2)
                If (LoLmt_NA_Flag = True And HiLmt_NA_Flag = True) Then TheExec.Flow.TestLimit MeasFreq, , , , , , unitHz, , Tname:=TestNameInput, ForceResults:=tlForceNone
                If (LoLmt_NA_Flag = True And HiLmt_NA_Flag = False) Then TheExec.Flow.TestLimit MeasFreq, , Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput, ForceResults:=tlForceNone
                If (LoLmt_NA_Flag = False And HiLmt_NA_Flag = True) Then TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, , , , , unitHz, , Tname:=TestNameInput, ForceResults:=tlForceNone
                If (LoLmt_NA_Flag = False And HiLmt_NA_Flag = False) Then TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput, ForceResults:=tlForceNone
                Call UpdateDLogColumns__False
            Else
                TestNameInput = TestNameInput + CStr(TestSeqNum)
                ''TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceNone
                If (LoLmt_NA_Flag = True And HiLmt_NA_Flag = True) Then TheExec.Flow.TestLimit MeasFreq, , , , , , unitHz, , Tname:=TestNameInput, ForceResults:=tlForceNone
                If (LoLmt_NA_Flag = True And HiLmt_NA_Flag = False) Then TheExec.Flow.TestLimit MeasFreq, , Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput, ForceResults:=tlForceNone
                If (LoLmt_NA_Flag = False And HiLmt_NA_Flag = True) Then TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, , , , , unitHz, , Tname:=TestNameInput, ForceResults:=tlForceNone
                If (LoLmt_NA_Flag = False And HiLmt_NA_Flag = False) Then TheExec.Flow.TestLimit MeasFreq, Freq_LowLimit, Freq_HighLimit, , , , unitHz, , Tname:=TestNameInput, ForceResults:=tlForceNone
            End If
        End If
    Else
        If Mid(TestLimitByPin_VFI, 2, 1) = "T" Then
            If IsDifferentialPin = True Then
                For p = 0 To MeasFreq.Pins.Count - 1 Step 2 ' freq counter result of differential pins is stored in positive pin
                     ''TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p + 1), unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum) + "_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), ForceResults:=tlForceFlow
                     TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p + 1), unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
                Next p
            Else
                For p = 0 To MeasFreq.Pins.Count - 1
                     '' TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p), unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum) + "_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), ForceResults:=tlForceFlow
                     TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p), unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
                Next p
            End If
        Else
            '' TheExec.Flow.TestLimit resultVal:=MeasFreq, unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum) + "_" + "@COND:PATTERN=" + PATT_ExculdePath(Pat), ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=MeasFreq, unit:=unitHz, Tname:=TestNameInput + CStr(TestSeqNum), ForceResults:=tlForceFlow
        End If
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151224 New for Mbist Shmoo
Public Function auto_storePower_perSite(power_pins As String, ByRef m_pwrStored As PinListData)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_storePower_perSite"

    Dim p_ary() As String
    Dim p_cnt As Long
    Dim i As Long
    Dim ss As Variant
    Dim m_pvalue As Double
    Dim p_pin As String
    
    ss = TheExec.Sites.SiteNumber
    TheExec.DataManager.DecomposePinList power_pins, p_ary, p_cnt
    
    Dim m_pin As Variant
    Dim m_match_flag As Boolean
    m_match_flag = False
    
    'Set m_pwrStored = New PinListData
    
    For i = 0 To p_cnt - 1
        p_pin = p_ary(i)
        For Each m_pin In m_pwrStored.Pins
            If (UCase(CStr(m_pin)) = UCase(p_pin)) Then
                m_match_flag = True
            End If
        Next m_pin
        
        If (m_match_flag = False) Then
            m_pwrStored.AddPin (p_pin)
        End If
        
        m_pwrStored.Pins(p_pin).Value(ss) = 0# ''''<MUST> initialize
        If TheExec.DataManager.ChannelType(p_pin) <> "N/C" Then
            m_pwrStored.Pins(p_pin).Value(ss) = TheHdw.DCVS.Pins(p_pin).Voltage.Main.Value
            m_pvalue = m_pwrStored.Pins(p_pin).Value(ss)
        End If
        'If (True) Then Debug.Print funcName + ":: Site(" + CStr(ss) + ") " + FormatNumeric(p_pin, 12) + " = " + Format(m_pvalue, "#0.######")
    Next i
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''This function should be called inside the Site-Loop
Public Function auto_getPowerValue_fromStore(inpwrpin As String, pwrGroup As String, pwrStore As PinListData) As Double

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_getPowerValue_fromStore"

    Dim p_ary() As String
    Dim p_cnt As Long
    Dim i As Long
    Dim m_pwrpin As String
    Dim ss As Variant
    
    ss = TheExec.Sites.SiteNumber
    auto_getPowerValue_fromStore = 0# ''''initialize

    TheExec.DataManager.DecomposePinList pwrGroup, p_ary, p_cnt
    
    For i = 0 To p_cnt - 1
        m_pwrpin = UCase(p_ary(i))
        If (m_pwrpin = UCase(inpwrpin)) Then
            If TheExec.DataManager.ChannelType(inpwrpin) <> "N/C" Then
                auto_getPowerValue_fromStore = pwrStore.Pins(inpwrpin).Value(ss)
            End If
            Exit For
        End If
        If (True) Then Debug.Print funcName + ":: Site(" + CStr(ss) + ") " + FormatNumeric(inpwrpin, 12) + " = " + Format(auto_getPowerValue_fromStore, "#0.######")
    Next i
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

