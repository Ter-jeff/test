Attribute VB_Name = "LIB_Harvest"
Option Explicit

Type Harvest_ALTBackup_Type
    'Pin As String
    mode As Long
    IfPowerPin As Boolean
    Voltage As Double
    Current As Double
    SrcCurrentRange As Double
    SourceFlodLimit As Double
    SinkFoldLimit As Double
    FilterValue As Double
End Type

Public gHarvDic_MappingTabble_SrcStr As New Dictionary
Public gHarvDic_MappingTabble_FlagContained As New Dictionary
Public gHarvDic_MappingTabble_BitNum As New Dictionary

Public Type ATE_STR_Summary_Table
    PTR_Test_Name As String
    operator As String
    value As String 'Input Flag
End Type

Public PTR_Flag() As ATE_STR_Summary_Table
Public ATE_STR_Summary_Table_Parse_Flag As Boolean
Public glb_pattern_name As String 'CJR:6/21/24



Public Function Harvest_StrExpand(ByVal InputStr As String, Optional ByVal JoinSymbol As String = ",", Optional ByRef DefaultStr As String, Optional ByRef StarIdx As Long, Optional ByRef EndIdx As Long, Optional ByRef HarvFuncName As String) As String
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_StrExpand"
    
    Dim i As Long
    Dim StepOrder As Long
    Dim OutputStrArr() As String
    
    Dim TempHarvCode() As String
    Dim TempHarvCodeAry() As String
    Dim HarvCodeIdx As Long
    
    
    If (InStr(LCase(InputStr), "(") <> 0) And (InStr(LCase(InputStr), ")") <> 0) Then
        'input : BitPos2Dec(b1000), BitPos2Dec(F_GFX_HARV[4:0])
        'return: b1000
        StarIdx = CLng(InStr(InputStr, "(") + 1)
        EndIdx = CLng(InStr(InputStr, ")"))
        HarvFuncName = mid(InputStr, 1, InStr(InputStr, "(") - 1)
        Harvest_StrExpand = mid(InputStr, StarIdx, (EndIdx - StarIdx))
    Else
        If (InStr(InputStr, "[") = 0) And (InStr(InputStr, ":") = 0) And (InStr(InputStr, "]") = 0) Or InStr(InputStr, "&") <> 0 Then
            Harvest_StrExpand = InputStr
        ElseIf (InStr(InputStr, "[") <> 0) And (InStr(InputStr, ":") = 0) And (InStr(InputStr, "]") <> 0) Then
            TempHarvCode = Split(InputStr, "[")
            TempHarvCodeAry = Split(TempHarvCode(UBound(TempHarvCode)), "]")
            HarvCodeIdx = CLng(TempHarvCodeAry(0))
            Harvest_StrExpand = TempHarvCode(0) & HarvCodeIdx
        ElseIf (InStr(InputStr, "[") = 0) Or (InStr(InputStr, ":") = 0) Or (InStr(InputStr, "]") = 0) Then
            TheExec.Datalog.WriteComment "<Error> Harvest_StrExpand: input might be wrong, please check."
        Else
            'input : F_GFX_HARV[4:0]
            'return: F_GFX_HARV4,F_GFX_HARV3,F_GFX_HARV2,F_GFX_HARV1,F_GFX_HARV0

            DefaultStr = mid(InputStr, 1, InStr(InputStr, "[") - 1)
            StarIdx = CLng(mid(InputStr, InStr(InputStr, "[") + 1, InStr(InputStr, ":") - InStr(InputStr, "[") - 1))
            EndIdx = CLng(mid(InputStr, InStr(InputStr, ":") + 1, InStr(InputStr, "]") - InStr(InputStr, ":") - 1))
            If StarIdx > EndIdx Then
                StepOrder = -1
            Else
                StepOrder = 1
            End If
            ReDim OutputStrArr(Abs(StarIdx - EndIdx))
            For i = 0 To UBound(OutputStrArr)
                OutputStrArr(i) = DefaultStr & CStr(StarIdx + i * StepOrder)
            Next i
            Harvest_StrExpand = Join(OutputStrArr, JoinSymbol)

        End If
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_StrExpand") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function Harvest_GetAllSiteFlagState(ByVal SiteFlag As String, Optional ByVal TrueVal As Long = 1, Optional ByVal FalseVal As Long = 0, Optional ByVal ClearVal As Long = -1) As SiteLong
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_GetAllSiteFlagState"
    
    Dim vsite As Variant
    Dim ReturnSLng As New SiteLong

    For Each vsite In TheExec.sites.Selected
        If TheExec.sites.item(vsite).FlagState(SiteFlag) = logicTrue Then
            ReturnSLng(vsite) = TrueVal
        ElseIf TheExec.sites.item(vsite).FlagState(SiteFlag) = logicFalse Then
            ReturnSLng(vsite) = FalseVal
        ElseIf TheExec.sites.item(vsite).FlagState(SiteFlag) = logicClear Then
            ReturnSLng(vsite) = ClearVal
            TheExec.Datalog.WriteComment "<Warning> Harvest_GetAllSiteFlagState : " & SiteFlag & "=logicClear, at site" & CStr(vsite)
        Else
        End If
    Next vsite
    Set Harvest_GetAllSiteFlagState = ReturnSLng
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_GetAllSiteFlagState") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_SetAllSiteFlagState(ByVal SiteFlag As String, ByVal FlagStateSL As SiteLong) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_SetAllSiteFlagState"
    Dim vsite As Variant
    Dim TempFlag() As String
    Dim i, j As Integer: i = 0: j = 0
    Dim SStr_TempAry As String
    Dim StarIdx, EndIdx As Long
    Dim StepOrder As Long
    Dim OutputStrArr() As String
    Dim StoreFlagIdx As Long: StoreFlagIdx = 0
    Dim DefaultStr As String
    
    For Each vsite In TheExec.sites.Selected
        If InStr(1, SiteFlag, "&") <> 0 Then
            TempFlag = Split(SiteFlag, "&")
            If InStr(1, SiteFlag, "[") <> 0 And InStr(1, SiteFlag, "]") Then
                For i = 0 To UBound(TempFlag)
                    DefaultStr = mid(TempFlag(i), 1, InStr(TempFlag(i), "[") - 1)
                    StarIdx = CLng(mid(TempFlag(i), InStr(TempFlag(i), "[") + 1, InStr(TempFlag(i), ":") - InStr(TempFlag(i), "[") - 1))
                    EndIdx = CLng(mid(TempFlag(i), InStr(TempFlag(i), ":") + 1, InStr(TempFlag(i), "]") - InStr(TempFlag(i), ":") - 1))
                    If StarIdx > EndIdx Then
                        StepOrder = -1
                    Else
                        StepOrder = 1
                    End If
                    ReDim Preserve OutputStrArr(Abs(StarIdx - EndIdx + StoreFlagIdx))
                    For j = StoreFlagIdx To UBound(OutputStrArr)
                        OutputStrArr(j) = DefaultStr & CStr(StarIdx + j * StepOrder + StoreFlagIdx)
                    Next j
                    StoreFlagIdx = UBound(OutputStrArr) + 1
                Next i
               
                SStr_TempAry = GlbUtility.Dec2Bin(CDbl(FlagStateSL), UBound(OutputStrArr) + 1)
                For j = 0 To UBound(OutputStrArr)
                    If mid(SStr_TempAry, j + 1, 1) = 0 Then
                        TheExec.sites.item(vsite).FlagState(OutputStrArr(j)) = logicFalse
                    Else
                        TheExec.sites.item(vsite).FlagState(OutputStrArr(j)) = logicTrue
                    End If
                Next j
                
            ElseIf FlagStateSL(vsite) = 8 Then
                TheExec.sites.item(vsite).FlagState(TempFlag(i)) = logicFalse
                TheExec.sites.item(vsite).FlagState(TempFlag(i + 1)) = logicTrue
            Else
                TheExec.sites.item(vsite).FlagState(TempFlag(i)) = logicTrue
                TheExec.sites.item(vsite).FlagState(TempFlag(i + 1)) = logicFalse
            End If
        Else
            If FlagStateSL(vsite) = logicTrue Or FlagStateSL(vsite) = logicFalse Or FlagStateSL(vsite) = logicClear Then
                TheExec.sites.item(vsite).FlagState(SiteFlag) = FlagStateSL(vsite)
                If TheExec.sites.item(vsite).FlagState(SiteFlag) = logicTrue Then
                    'Record this fail flag to disable compare
                    If Not glb_Harvest_Fail_Flag_From_eFuseRead_Dict.Exists((vsite & SiteFlag)) Then
                        glb_Harvest_Fail_Flag_From_eFuseRead_Dict.Add (vsite & SiteFlag), "TRUE"
                    End If
                End If
            Else
                TheExec.Datalog.WriteComment "<Error> Harvest_SetAllSiteFlagState: do not support this case."
            End If
        End If
    Next vsite

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_SetAllSiteFlagState") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231124][All][Tank] Check sheet exist or not
Public Function Harvest_DigSrc_MappingTableParsing() As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funcName As String:: funcName = "Harvest_DigSrc_MappingTableParsing"
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim sheetName As String
    Dim ws As Worksheet
    Dim MaxRow As Long
    Dim MaxColumn As Long
    Dim vTableContent As Variant
    Dim vDigSrcInfo() As Variant
    Dim HeaderRowArr() As String
    Dim HeaderCnt As Long
    Dim PatValue As String
    Dim patArr() As String
    Dim PatArrCnt As Long
    Dim PatDic As New Dictionary
    Dim BitValue As String
    Dim BitArr() As Long
    Dim BitArrCnt As Long
    Dim BitArr_StartRow() As Long
    Dim BitArr_EndRow() As Long
    Dim HarvDSSC_Value As String
    Dim HarvDSSC_StrArr() As String
    Dim HarvDSSC_FlagDic As New Dictionary
    Dim HarvDSSC_FlagArr() As String
    Dim HarvDSSC_FlagArrCnt As Long
    Dim HarvDicKey  As String
    
    Dim TempCnt As Long
    Dim TempStrArr() As String
    
    Const DefaultSheetName = "HARVMappingTable_"
    Const Col1_Pat = 1
    Const Col2_Bit = 2
    Const Col4_DSSC = 4
    Application.ScreenUpdating = False
    
    Static bIsMappingTableParsed As Boolean
    If bIsMappingTableParsed = False Then
        gHarvDic_MappingTabble_SrcStr.RemoveAll
        gHarvDic_MappingTabble_FlagContained.RemoveAll
        gHarvDic_MappingTabble_BitNum.RemoveAll
    
        If TheExec.TesterMode = testModeOffline Then
            sheetName = DefaultSheetName & "CP1"
        Else
            sheetName = DefaultSheetName & UCase(currentJobName)
        End If

        If GetSheetInfo(sheetName, MaxRow, MaxColumn, vDigSrcInfo) Then
            ''' need to transpose for this mapping table
            vDigSrcInfo = Application.WorksheetFunction.Transpose(vDigSrcInfo)
            'record the header location
            HeaderCnt = 0
            For i = 1 To MaxRow
                If UCase(vDigSrcInfo(Col1_Pat, i)) = UCase("PatternName") Then
                    ReDim Preserve HeaderRowArr(HeaderCnt)
                    ReDim Preserve BitArr_StartRow(HeaderCnt)
                    ReDim Preserve BitArr_EndRow(HeaderCnt)
                    HeaderRowArr(HeaderCnt) = i
                    BitArr_StartRow(HeaderCnt) = i + 1
                    If HeaderCnt <> 0 Then
                        BitArr_EndRow(HeaderCnt - 1) = i - 1
                    End If
                    HeaderCnt = HeaderCnt + 1
                End If
            Next i
            BitArr_EndRow(HeaderCnt - 1) = i - 1
            
            For i = 0 To UBound(HeaderRowArr)
                BitArrCnt = 0
                PatArrCnt = 0
                For j = BitArr_StartRow(i) To BitArr_EndRow(i)
                    'only consider from 0 to XXX, from small to big, LSB to MSB
                    BitValue = vDigSrcInfo(Col2_Bit, j)
                    If BitValue <> "" Then
                        If IsNumeric(BitValue) Then
                            ReDim Preserve BitArr(BitArrCnt)
                            BitArr(BitArrCnt) = j
                        Else
                            TempStrArr = Split(BitValue, "-")
                            TempCnt = Abs(TempStrArr(0) - TempStrArr(1))
                            ReDim Preserve BitArr(BitArrCnt + TempCnt)
                            For k = BitArrCnt To UBound(BitArr)
                                BitArr(k) = j
                            Next k
                            BitArrCnt = BitArrCnt + TempCnt
                        End If
                        BitArrCnt = BitArrCnt + 1
                    End If
                    
                    'record pattern
                    PatValue = UCase(vDigSrcInfo(Col1_Pat, j))
                    If PatValue <> "" Then
                        If PatDic.Exists(PatValue) Then
                            TheExec.Datalog.WriteComment "<Error> Harvest_DigSrc_MappingTableParsing : duplicate Pat Name."
                            Exit Function
                        End If
                        PatDic.Add PatValue, vbNullString
                        ReDim Preserve patArr(PatArrCnt)
                        patArr(PatArrCnt) = PatValue
                        PatArrCnt = PatArrCnt + 1
                    End If
                Next j
                
                For j = Col4_DSSC To MaxColumn
                    If vDigSrcInfo(j, HeaderRowArr(i)) <> "" Then
                        ReDim HarvDSSC_StrArr(UBound(BitArr))
                        ReDim HarvDSSC_FlagArr(0)
                        HarvDSSC_FlagArrCnt = 0
                        HarvDSSC_FlagDic.RemoveAll
                        For k = 0 To UBound(BitArr)
                            HarvDSSC_Value = vDigSrcInfo(j, BitArr(k))
                            HarvDSSC_StrArr(k) = HarvDSSC_Value
                            If IsNumeric(HarvDSSC_Value) = False Then
                                HarvDSSC_Value = Replace(HarvDSSC_Value, "!", "")
                                If HarvDSSC_FlagDic.Exists(HarvDSSC_Value) = False Then
                                    HarvDSSC_FlagDic.Add HarvDSSC_Value, vbNullString
                                    ReDim Preserve HarvDSSC_FlagArr(HarvDSSC_FlagArrCnt)
                                    HarvDSSC_FlagArr(HarvDSSC_FlagArrCnt) = HarvDSSC_Value
                                    HarvDSSC_FlagArrCnt = HarvDSSC_FlagArrCnt + 1
                                End If
                            End If
                        Next k
                        For k = 0 To UBound(patArr)
                            'PatArr, HarvResult_StrArr, HarvResult_FlagArr are all UCASE
                            HarvDicKey = UCase(patArr(k) & "_" & vDigSrcInfo(j, HeaderRowArr(i)))
                            gHarvDic_MappingTabble_SrcStr.Add HarvDicKey, UCase(Join(HarvDSSC_StrArr, ","))
                            gHarvDic_MappingTabble_FlagContained.Add HarvDicKey, UCase(Join(HarvDSSC_FlagArr, ","))
                            gHarvDic_MappingTabble_BitNum.Add HarvDicKey, UBound(BitArr) + 1
                        Next k
                    End If
                Next j
            Next i
            
            bIsMappingTableParsed = True
        End If
    Else
        TheExec.Datalog.WriteComment "There is no " & sheetName & " Sheet in this workbook"
    End If
    Application.ScreenUpdating = True

Exit Function 'Add errHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_DigSrc_MappingTableParsing") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20240503][All][Clyde] Fixed Harvest can not get Source
Public Function Harvest_CreateDigSrc(ByVal HarvPat As String, _
                                     Optional ByVal HarvDSSCHeader_atTable As String = "Value", _
                                     Optional ByRef DigSrcDsp As DSPWave, _
                                     Optional ByRef DigSrcDsp_SampleSize As Long, _
                                     Optional debugF As Boolean = False) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_CreateDigSrc"

    Dim vsite As Variant
    Dim i As Long
    Dim HarvDicKey As String
    Dim OutputDSP As New DSPWave
    Dim OutputSrcStr_SiteVar As New SiteVariant
    Dim SrcStr_FromTable As String
    Dim FlagContained_FromTable As String
    Dim FlagContainedArr() As String
    Dim tempFlagState As New SiteLong
    Dim tempStr As String
    Dim patcnt As Long
    Dim PatArr_absolute() As String

    HarvDicKey = UCase(HarvPat & "_" & HarvDSSCHeader_atTable)
    If gHarvDic_MappingTabble_SrcStr.Exists(HarvDicKey) = False Then
        TheExec.Datalog.WriteComment "<Error> Harvest_CreateDigSrc : Input Pat and Header are not in the mapping table."
    Else
        SrcStr_FromTable = gHarvDic_MappingTabble_SrcStr(UCase(HarvDicKey))
        FlagContained_FromTable = gHarvDic_MappingTabble_FlagContained(UCase(HarvDicKey))
        DigSrcDsp_SampleSize = gHarvDic_MappingTabble_BitNum(UCase(HarvDicKey))
        
        OutputSrcStr_SiteVar = SrcStr_FromTable
        FlagContainedArr = Split(FlagContained_FromTable, ",")
        For i = 0 To UBound(FlagContainedArr)
            tempFlagState = Harvest_GetAllSiteFlagState(FlagContainedArr(i), 1, 0, 1)
            For Each vsite In TheExec.sites.Selected
                OutputSrcStr_SiteVar = Replace(OutputSrcStr_SiteVar, UCase(FlagContainedArr(i)), CStr(tempFlagState))
            Next vsite
        Next i
        For Each vsite In TheExec.sites.Selected
            OutputSrcStr_SiteVar = Replace(OutputSrcStr_SiteVar, "!1", "0")
            OutputSrcStr_SiteVar = Replace(OutputSrcStr_SiteVar, "!0", "1")
        Next vsite
        
        OutputDSP.CreateConstant 0, DigSrcDsp_SampleSize, DspLong
        Harvest_CreateDSPbySiteVar OutputSrcStr_SiteVar, OutputDSP
        
        If debugF Then
            TheExec.Datalog.WriteComment "----Harvest DigSrc"
            TheExec.Datalog.WriteComment UCase(HarvPat)
            TheExec.Datalog.WriteComment "Src Info at Table : " & SrcStr_FromTable
            For Each vsite In TheExec.sites.Selected
                TheExec.Datalog.WriteComment OutputSrcStr_SiteVar
            Next vsite
            TheExec.Datalog.WriteComment "----Harvest DigSrc End"
        End If
        
        For Each vsite In TheExec.sites.Selected
            DigSrcDsp = OutputDSP.COPY
        Next vsite
        
        'Print Info
        Call GetPatsFromPatSets(HarvPat, PatArr_absolute(), patcnt, True)
        For Each vsite In TheExec.sites.Selected
            tempStr = "Site" & vsite & " DigSrc pattern = " & PatArr_absolute(0) & ", HarvestHeader at Table = " & HarvDSSCHeader_atTable & ", Src Bits = " & CStr(DigSrcDsp_SampleSize) & ", HarvestSourceCode [ First(L) ==> Last(R) ] : "
            For i = 0 To DigSrcDsp_SampleSize - 1
                tempStr = tempStr & OutputDSP.Element(i)
            Next i
            Call ShowLog(tempStr)
        Next vsite
    End If
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Harvest", funcName) 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Harvest_CreateDSPbySiteVar(ByVal InputSiteVar As SiteVariant, ByRef OutputDSP As DSPWave) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_CreateDSPbySiteVar"

    Dim vsite As Variant
    Dim i As Long
    Dim TempArr() As String
    
    For Each vsite In TheExec.sites.Selected
        TempArr = Split(InputSiteVar, ",")
        For i = 0 To UBound(TempArr)
            OutputDSP.Element(i) = CLng(TempArr(i))
        Next i
    Next vsite

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_CreateDSPbySiteVar") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function PrintDic(ByVal InputDicName As String, ByVal InputDic As Dictionary) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "PrintDic"

    Dim vsite As Variant
    Dim i As Long
    Dim j As Long
    
    Dim strTypeName As String
    Dim tempPLD As New PinListData
    Dim TempDSP As New DSPWave
    
    Const bDatalogWrite = True
    Const bDebugPrint = False
    Const bOutputWindow = False
    
    If InputDic.Count = 0 Then
        Call ShowLog(InputDicName & " is Empty.", bDatalogWrite, bDebugPrint, bOutputWindow)
        Exit Function
    Else
        strTypeName = TypeName(InputDic(InputDic.Keys(0)))
    End If
    
    Call ShowLog("----Print Dictionary : " & InputDicName, bDatalogWrite, bDebugPrint, bOutputWindow)
    For i = 0 To InputDic.Count - 1
        If strTypeName = "Long" Or strTypeName = "Double" Or strTypeName = "String" Then
            '(Boolean, Variant did not try)
            Call ShowLog("Key:" & InputDic.Keys(i) & ", Item:" & CStr(InputDic(InputDic.Keys(i))), bDatalogWrite, bDebugPrint, bOutputWindow)
        
        ElseIf InStr(strTypeName, "ISite") Then
            'ISiteLong, ISiteDouble, (ISiteBoolean, ISiteVariant did not try)
            Call ShowLog("Key:" & InputDic.Keys(i), bDatalogWrite, bDebugPrint, bOutputWindow)
            For Each vsite In TheExec.sites.Selected
                Call ShowLog(vbTab & "Site" & vsite & ", " & CStr(InputDic(InputDic.Keys(i))), bDatalogWrite, bDebugPrint, bOutputWindow)
            Next vsite
        
        ElseIf strTypeName = "IPinListData" Then
            Call ShowLog("Key:" & InputDic.Keys(i), bDatalogWrite, bDebugPrint, bOutputWindow)
            tempPLD = InputDic(InputDic.Keys(i))
            For j = 0 To tempPLD.pins.Count - 1
                Call ShowLog(vbTab & "Pin=" & tempPLD.pins(j), bDatalogWrite, bDebugPrint, bOutputWindow)
                For Each vsite In TheExec.sites.Selected
                    Call ShowLog(vbTab & vbTab & "Site" & vsite & ", " & CStr(tempPLD.pins(j).value(vsite)), bDatalogWrite, bDebugPrint, bOutputWindow)
                Next vsite
            Next j
        
        ElseIf strTypeName = "IDspWave_i" Then
            Call ShowLog("Key:" & InputDic.Keys(i), bDatalogWrite, bDebugPrint, bOutputWindow)
            TempDSP = InputDic(InputDic.Keys(i))
            'add later
        
        Else
            Call ShowLog("did not support yet", bDatalogWrite, bDebugPrint, bOutputWindow)
        
        End If
    Next i
    Call ShowLog("----Print Dictionary End, Count = " & CStr(i), bDatalogWrite, bDebugPrint, bOutputWindow)
            
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "PrintDic") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function ShowLog(ByVal InputStr As String, _
                        Optional ByVal bDatalogWrite As Boolean = True, _
                        Optional ByVal bDebugPrint As Boolean = False, _
                        Optional ByVal bOutputWindow As Boolean = False) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "PrintDic"
    
    If bDatalogWrite Then
        TheExec.Datalog.WriteComment InputStr
    End If
    
    If bDebugPrint Then
        Debug.Print InputStr
    End If
    
    If bOutputWindow Then
        TheExec.AddOutput InputStr
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "ShowLog") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



' This function sets up scan data capture at the Start of Body.
' Public Function StartOfBodyIPF(argc As Integer, argv() As String)
Public Function Harvest_CMEM_InitSetup(Optional ByVal CapSize As Long = -1) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_CMEM_InitSetup"
    
    If TheExec.TesterMode = testModeOnline Then
        Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone) ' Resets CMEM
        If glb_TesterType = "UltraFLEXplus" Then
            ' Capture all failures (full 256K single or 512K dual).
    '        Call TheHdw.Digital.CMEM.SetCaptureConfig(CapSize, CmemCaptFail)
            Call TheHdw.Digital.CMEM.SetCaptureConfig(16777216, CmemCaptFail, tlCMEMCaptureSource_PatPassFailData)
            TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + _
                                                tlCMEMVMVectorOffset + _
                                                tlCMEMModCycle
            TheHdw.Digital.CMEM.CaptureLimitMode = tlDigitalCMEMCaptureLimitMode_EnableResetOnModule
            TheHdw.Digital.CMEM.CaptureLimit = 512
        Else
            Call TheHdw.Digital.CMEM.SetCaptureConfig(CapSize, CmemCaptFail)
            TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + _
                                                tlCMEMVMVectorOffset + _
                                                tlCMEMModCycle
        End If
    End If
        
    If glb_TesterType = "UltraFLEXplus" Then
        TheHdw.Digital.Patgen.ScanBurstEnabled = True
        TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = True
    End If
        
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_CMEM_InitSetup") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' This function turns off central capture at the End of Body.
' Public Function EndOfBodyIPF(argc As Integer, argv() As String)
Public Function Harvest_CMEM_Stop(ConditionArr() As String, CustomCondition() As Boolean, failFlagArr() As String, Optional initPatBool As Boolean) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_CMEM_Stop"
    
    Dim i As Integer
    Dim TempPinGupStr() As String
    Dim site As Variant
    
    If TheExec.TesterMode = testModeOnline Then
        ' Use this statement to turn off the Pin CMEM capture.
        Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone) ' Resets CMEM
        ' Use this statement to turn off the Central CMEM capture.
        TheHdw.Digital.CMEM.CentralFields = tlCMEMNone
    End If
     
    If initPatBool = False Then
        For i = 0 To UBound(CustomCondition)
            TempPinGupStr = Split(ConditionArr(i), ":")
            If CustomCondition(i) = True Then
                For Each site In TheExec.sites
                    If TheExec.sites.item(site).FlagState(failFlagArr(i)) = logicTrue Then
                        TheHdw.Digital.pins(TempPinGupStr(1)).DisableCompare = False
                        TheExec.Datalog.WriteComment "--(Disable Pin mask Feature)--" & "site = " & site
                    End If
                Next site
            End If
        Next i
    End If
    
     ''''' New request for pattern pin group '''''
    If glb_TesterType = "UltraFLEXplus" Then
        TheHdw.Digital.Patgen.ScanBurstEnabled = False
        TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = False
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_CMEM_Stop") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_FailFlagSplit(ByVal Condition_and_FailFlag As String, _
                                      ByRef ConditionArr() As String, ByRef failFlagArr() As String, ByRef CustomCondition() As Boolean) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_FailFlagSplit"
    
    'for temp format test
    'Condition_and_FailFlag = P:ECPU_CORE0(F_ECPU_CORE0);P:ECPU_CORE1(F_ECPU_CORE1)
    Dim i As Long
    Dim tempStrArr0() As String
    Dim tempStrArr1() As String
    Dim tempStrArr2() As String
    
    Dim tempCondition As String
    'Dim ConditionArr() As String
    Dim tempFailFlag As String
    'Dim FailFlagArr() As String
    
    'If Condition_and_FailFlag <> "" Then
    tempStrArr0 = Split(Condition_and_FailFlag, ";")
    ReDim CustomCondition(UBound(tempStrArr0))
    For i = 0 To UBound(tempStrArr0)
        If InStr(tempStrArr0(i), "(") > 0 Then
            tempStrArr1 = Split(tempStrArr0(i), "(")
            tempCondition = tempStrArr1(0)
            tempStrArr2 = Split(tempStrArr1(1), ")")
            tempFailFlag = tempStrArr2(0)
            If InStr(1, UCase(tempStrArr0(i)), UCase("DisableCompare")) <> 0 Then
                CustomCondition(i) = True
            End If
            ReDim Preserve ConditionArr(i)
            ReDim Preserve failFlagArr(i)
            ConditionArr(i) = tempCondition
            failFlagArr(i) = tempFailFlag
        End If
    Next i

'        For i = 0 To UBound(FailFlagArr)
'            Call Harvest_Decision(FailFlagArr(i), ConditionArr(i))
'        Next i
    'End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_FailFlagSplit") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function Harvest_Decision(ByVal Harv_FailFlag As String, Optional ByVal Harv_InputCondition As String = vbNullString, Optional PatName As String, Optional CustomCondition As Boolean) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_Decision"
    ''Harv_FailFlag = F_Gfx_Core0
    ''Harv_Condition = P:JTAG_TDO:323,325,330;V:EUSB_P1:0.05;F:EUSB_P2:3000000;I:EUSB_P1:0~0.05
    
    Const Cond_Case = 0
    Const Cond_Pin = 1

    Dim vsite As Variant
    Dim sInstName As String
    
    Dim i As Long
    Dim str_InputCondArr() As String
    Dim str_CondArr() As String
    Dim str_TempArr() As String
    Dim str_temp As String
    
    Dim str_MeasPin As String
    Dim str_PatVectorFilter As String
    Dim str_PassRange As String
    Dim str_ForceVal As String
    Dim str_MeterCurrRange As String
    Dim str_FreqInterval As String
    Dim str_ForceVal_Z As String
    Dim str_MeterCurrRange_Z As String
    Dim Sdbl_MeasVal As New SiteDouble
    Dim Sbln_FailResult As New SiteBoolean
    Dim Sbln_PatternPass As New SiteBoolean
    Dim Sbln_PatternResult As New SiteLong
    
    If (Harv_InputCondition <> "") And (Harv_FailFlag <> "") Then
        Sbln_FailResult = False
        str_InputCondArr = Split(Harv_InputCondition, ";")
        'for measure, only implement single pin first
        For i = 0 To UBound(str_InputCondArr)
            str_CondArr = Split(str_InputCondArr(i), ":")
            If (UCase(str_CondArr(Cond_Case)) = "P") Or (UBound(str_CondArr) = 0) Then
                'P:JTAG_TDO:323,330
                If UBound(str_CondArr) = 0 Then
                    str_MeasPin = str_CondArr(0)
                Else
                    str_MeasPin = str_CondArr(Cond_Pin)
                End If
                If UBound(str_CondArr) >= 2 Then
                    str_PatVectorFilter = str_CondArr(2)
                End If
                If glb_TesterType = "UltraFLEXplus" Then
                    Call Harvest_CMEM_PostResult_UP(Harv_FailFlag, str_MeasPin, , PatName, CustomCondition)
                Else
                    Call Harvest_CMEM_PostResult_UF(Harv_FailFlag, str_MeasPin)
                End If

                
            ElseIf UCase(str_CondArr(Cond_Case)) = "V" Then
                'V:AOP_FUNC0:0~0.75:0.02;
                str_MeasPin = str_CondArr(Cond_Pin)
                str_PassRange = str_CondArr(2)
                str_ForceVal = str_CondArr(3)
                Call Harvest_MeasureVolt(Sdbl_MeasVal, str_MeasPin, str_ForceVal)
                Call Harvest_FailCheck(Sbln_FailResult, Sdbl_MeasVal, str_PassRange)
                
            ElseIf UCase(str_CondArr(Cond_Case)) = "I" Then
                'I:AOP_FUNC0:0~0.02:0.75|0.05
                str_MeasPin = str_CondArr(Cond_Pin)
                str_PassRange = str_CondArr(2)
                str_TempArr = Split(str_CondArr(3), "|")
                str_ForceVal = str_TempArr(0)
                str_MeterCurrRange = str_TempArr(1)
                Call Harvest_MeasureCurr(Sdbl_MeasVal, str_MeasPin, str_ForceVal, str_MeterCurrRange)
                Call Harvest_FailCheck(Sbln_FailResult, Sdbl_MeasVal, str_PassRange)
                
            ElseIf UCase(str_CondArr(Cond_Case)) = "F" Then
                'F:EUSB_P2:0~3000000:0.01
                str_MeasPin = str_CondArr(Cond_Pin)
                str_PassRange = str_CondArr(2)
                str_FreqInterval = str_CondArr(3)
                Call Harvest_MeasureFreq(Sdbl_MeasVal, str_MeasPin)
                Call Harvest_FailCheck(Sbln_FailResult, Sdbl_MeasVal, str_PassRange)
            
            ElseIf UCase(str_CondArr(Cond_Case)) = "R" Then
                'R:LPDPRX_RX_D0_N:50~60:0.125|0.0078125
                str_MeasPin = str_CondArr(Cond_Pin)
                str_PassRange = str_CondArr(2)
                str_TempArr = Split(str_CondArr(3), "|")
                str_ForceVal = str_TempArr(0)
                str_MeterCurrRange = str_TempArr(1)
                Call Harvest_MeasureR(Sdbl_MeasVal, str_MeasPin, str_ForceVal, str_MeterCurrRange)
                Call Harvest_FailCheck(Sbln_FailResult, Sdbl_MeasVal, str_PassRange)
                
            ElseIf UCase(str_CondArr(Cond_Case)) = "Z" Then
                'Z:LPDPRX_RX_D0_N:50~60:0.125&0.250|0.0078125&0.0078125
                str_MeasPin = str_CondArr(Cond_Pin)
                str_PassRange = str_CondArr(2)
                str_TempArr = Split(str_CondArr(3), "|")
                str_ForceVal = Split(str_TempArr(0), "&")(0)
                str_MeterCurrRange = Split(str_TempArr(1), "&")(0)
                str_ForceVal_Z = Split(str_TempArr(0), "&")(1)
                str_MeterCurrRange_Z = Split(str_TempArr(1), "&")(1)
                Call Harvest_MeasureZ(Sdbl_MeasVal, str_MeasPin, str_ForceVal, str_MeterCurrRange, str_ForceVal_Z, str_MeterCurrRange_Z)
                Call Harvest_FailCheck(Sbln_FailResult, Sdbl_MeasVal, str_PassRange)
                
            Else
                TheExec.Datalog.WriteComment funcName & " did not consider this case, please check!"
            End If
            
            'can add site select for TTR
            'datalog did not add yet
        Next i
    ElseIf (Harv_FailFlag <> "") Then
        'Harv_InputCondition = "" => if pat fail, then F_flag = True
        Sbln_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        For Each vsite In TheExec.sites.Selected
            If Sbln_PatternPass(vsite) = False Then
                TheExec.sites.item(vsite).FlagState(Harv_FailFlag) = logicTrue
            Else
                If TheExec.sites.item(vsite).FlagState(Harv_FailFlag) = logicClear Then
                    TheExec.sites.item(vsite).FlagState(Harv_FailFlag) = logicFalse
                ElseIf TheExec.sites.item(vsite).FlagState(Harv_FailFlag) = logicTrue Then
                    'need check?
                End If
            End If
            Sbln_PatternResult = TheExec.sites.item(vsite).FlagState(Harv_FailFlag)
        Next vsite
        
'''        sInstName = theexec.DataManager.instanceName
        If UCase(Harv_FailFlag) Like "*SOC*" Or UCase(Harv_FailFlag) Like "*ECPU*" Then
            sInstName = TheExec.DataManager.instancename
            TheExec.flow.TestLimit Sbln_PatternResult, 0, 0, Tname:=sInstName & "_" & Harv_FailFlag
        End If
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_Decision") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_FailCheck(ByRef FailResult As SiteBoolean, ByVal MeasVal As SiteDouble, ByVal PassRange As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_FailCheck"

    Dim str_TempArr() As String
    Dim dbl_HiLimit As Double
    Dim dbl_LoLimit As Double
    
    str_TempArr = Split(PassRange, "~")
    dbl_LoLimit = CDbl(str_TempArr(0))
    dbl_HiLimit = CDbl(str_TempArr(1))
    
    FailResult = FailResult.LogicalOr(MeasVal.compare(LessThan, dbl_LoLimit))
    FailResult = FailResult.LogicalOr(MeasVal.compare(GreaterThan, dbl_HiLimit))

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_FailCheck") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureVolt(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceI As String, Optional ByVal WaitTime As String = "0.01") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureVolt"

    Dim str_PinInstrument As String
    Dim str_PinType As String
    Dim str_MeasPinH As String
    Dim str_MeasPinL As String
    
    str_PinInstrument = UCase(SortPinInstrument(MeasPin))
    If (str_PinInstrument = "DC-07") Or (str_PinInstrument = "DC-30") Or (str_PinInstrument = "DC-75") Then
        str_PinType = SortPinChannelType(MeasPin)
        If str_PinType = "DCDiffMeter" Then
            Call UVI80_DIFFMETER_INIT(MeasPin, str_MeasPinH, str_MeasPinL)
            Call Harvest_MeasureVolt_UVI80Diff(MeasVal, str_MeasPinH, str_MeasPinL)
        Else
            Call Harvest_MeasureVolt_UVI80(MeasVal, MeasPin, ForceI, WaitTime)
        End If
    ElseIf (str_PinInstrument = "HSD-U") Then
        Call Harvest_MeasureVolt_PPMU(MeasVal, MeasPin, ForceI, WaitTime)
    Else
        TheExec.Datalog.WriteComment funcName & " did not consider this case, please check!"
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureVolt") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureVolt_PPMU(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceI As String, Optional ByVal WaitTime As String = "0.01") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureVolt_PPMU"
    
    TheHdw.Digital.pins(MeasPin).Disconnect
    With TheHdw.PPMU.pins(MeasPin)
        .Gate = tlOff
        .ForceI CDbl(ForceI), CDbl(ForceI)
        .Connect
        .Gate = tlOn
    End With
    
    TheHdw.Wait (WaitTime)
    MeasVal = TheHdw.PPMU.pins(MeasPin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    With TheHdw.PPMU.pins(MeasPin)
        .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Disconnect
        .Gate = tlOff
    End With
    TheHdw.Wait (0.1 * ms)
    TheHdw.Digital.pins(MeasPin).Connect
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureVolt_PPMU") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureVolt_UVI80(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceI As String, Optional ByVal WaitTime As String = "0.01") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureVolt_UVI80"

    Dim typ_SaveCond As Harvest_ALTBackup_Type
    
    If TheHdw.DCVI.pins(MeasPin).Gate = True Then
        With TheHdw.DCVI.pins(MeasPin)
            typ_SaveCond.IfPowerPin = True
            'typ_SaveCond.Mode = .Mode
            'typ_SaveCond.Voltage = CDbl(FormatNumber(.Voltage, 3))
            typ_SaveCond.Current = CDbl(FormatNumber(.Current, 3))
            typ_SaveCond.SrcCurrentRange = CDbl(FormatNumber(.CurrentRange.value, 3))
        End With
    End If
    
    With TheHdw.DCVI.pins(MeasPin)
        If (CDbl(ForceI) = 0) And TheHdw.DCVI.pins(MeasPin).Gate <> True Then
            .mode = tlDCVIModeHighImpedance
            .Voltage = pc_Def_VFI_UVI80_VoltCalmp
            .BleederResistor = tlDCVIBleederResistorOff
            .Current = 0
            .Connect tlDCVIConnectHighSense
        Else
            .mode = tlDCVIModeCurrent
            .Voltage = pc_Def_VFI_UVI80_VoltCalmp
            .Current = CDbl(ForceI)
            .Connect tlDCVIConnectDefault
        End If
        If CDbl(ForceI) <> 0 Then
            .Gate(tlDCVIGateHiZ) = False
        End If
        .Gate = True
        .Meter.mode = tlDCVIMeterVoltage
    End With

    TheHdw.Wait CDbl(WaitTime)
    MeasVal = TheHdw.DCVI.pins(MeasPin).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)

    With TheHdw.DCVI.pins(MeasPin)
        If typ_SaveCond.IfPowerPin = True Then
            '.Mode = typ_SaveCond.Mode
            '.Voltage = typ_SaveCond.Voltage
            .CurrentRange.value = typ_SaveCond.SrcCurrentRange
            .Current = typ_SaveCond.Current
        Else
            If CDbl(.Voltage) <> 0 Then
                'TheExec.Datalog.WriteComment " *** UVI80 Force Conditon Rollback ***  " & "Gate: ON -> HiZOFF -> OFF"
                .Gate(tlDCVIGateHiZ) = False
                .Gate = False
            End If
            .Current = 0
            .Voltage = 0
            .Gate(tlDCVIGateHiZ) = False
            .BleederResistor = tlDCVIBleederResistorAuto
            .Disconnect
            .mode = tlDCVIModeCurrent
        End If
    End With
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureVolt_UVI80") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureVolt_UVI80Diff(ByRef MeasVal As SiteDouble, ByVal MeasPin_H As String, ByVal MeasPin_L As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureVolt_UVI80Diff"
        
    Call UVI80_DIFFMETER_SETUP(MeasPin_H, MeasPin_L)
    Call UVI80_DCDIFFMETER(MeasPin_H, MeasVal)
    Call UVI80_DIFFMETER_RELEASE(MeasPin_H, MeasPin_L)

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureVolt_UVI80Diff") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureCurr(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceV As String, ByVal MeterCurrRange As String, Optional ByVal WaitTime As String = "0.1") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureCurr"
    
    Dim str_PinInstrument As String
    
    str_PinInstrument = UCase(SortPinInstrument(MeasPin))
    If (str_PinInstrument = "HEXVS") Or (str_PinInstrument = "VHDVS") Or (str_PinInstrument = "VSM") Then
        Call Harvest_MeasureCurr_DCVS(MeasVal, MeasPin, ForceV, MeterCurrRange, WaitTime)
    ElseIf (str_PinInstrument = "HSD-U") Then
        Call Harvest_MeasureCurr_PPMU(MeasVal, MeasPin, ForceV, MeterCurrRange, WaitTime)
    ElseIf (str_PinInstrument = "DC-07") Then
        Call Harvest_MeasureCurr_UVI80(MeasVal, MeasPin, ForceV, MeterCurrRange, WaitTime)
    Else
        TheExec.Datalog.WriteComment funcName & " did not consider this case, please check!"
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureCurr") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureCurr_DCVS(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceV As String, ByVal MeterCurrRange As String, Optional ByVal WaitTime As String = "0.1") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureCurr_DCVS"
    
    Dim typ_SaveCond As Harvest_ALTBackup_Type
    Dim str_PinInstrument As String
    
    str_PinInstrument = UCase(SortPinInstrument(MeasPin))
    
    typ_SaveCond.Voltage = CDbl(FormatNumber(TheHdw.DCVS.pins(MeasPin).Voltage.Main, 3))
    typ_SaveCond.SourceFlodLimit = CDbl(FormatNumber(TheHdw.DCVS.pins(MeasPin).CurrentLimit.Source.FoldLimit.level.value, 3))
    typ_SaveCond.SinkFoldLimit = CDbl(FormatNumber(TheHdw.DCVS.pins(MeasPin).CurrentLimit.Sink.FoldLimit.level.value, 3))
    If str_PinInstrument = "VHDVS" Then
        typ_SaveCond.FilterValue = CDbl(FormatNumber(TheHdw.DCVS.pins(MeasPin).Meter.Filter.value, 3))
    End If
    typ_SaveCond.SrcCurrentRange = CDbl(FormatNumber(TheHdw.DCVS.pins(MeasPin).CurrentRange.value, 3))
    
    With TheHdw.DCVS.pins(MeasPin)
        .Voltage.Main = ForceV
        .Meter.mode = tlDCVSMeterCurrent
        If MeterCurrRange <> "0" Then
            .SetCurrentRanges CDbl(MeterCurrRange), CDbl(MeterCurrRange)
        End If
        .Gate = True
    End With
    
    TheHdw.Wait CDbl(WaitTime) 'need to care the settle time, did not consider yet
    MeasVal = TheHdw.DCVS.pins(MeasPin).Meter.Read(tlStrobe, 10)
        
    TheHdw.DCVS.pins(MeasPin).CurrentRange.value = typ_SaveCond.SrcCurrentRange 'Move here for avoiding Flodlimit Alarm - Carter, 20190507
    TheHdw.DCVS.pins(MeasPin).CurrentLimit.Source.FoldLimit.level.value = typ_SaveCond.SourceFlodLimit
    If str_PinInstrument <> VSM Then 'Skip switch sink fold limit for VSM to avoid error - JC-Chop MQHu, 20200205
        TheHdw.DCVS.pins(MeasPin).CurrentLimit.Sink.FoldLimit.level.value = typ_SaveCond.SinkFoldLimit
    End If
    If str_PinInstrument = "VHDVS" Then
        TheHdw.DCVS.pins(MeasPin).Meter.Filter.value = typ_SaveCond.FilterValue
    End If
    TheHdw.DCVS.pins(MeasPin).Voltage.Main = typ_SaveCond.Voltage

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureCurr_DCVS") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureCurr_PPMU(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceV As String, ByVal MeterCurrRange As String, Optional ByVal WaitTime As String = "0.1") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureCurr_DCVS"

    TheHdw.Digital.pins(MeasPin).Disconnect

    With TheHdw.PPMU.pins(MeasPin)
        .Gate = tlOff
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        .ForceV CDbl(ForceV), CDbl(MeterCurrRange)
        .Connect
        .Gate = tlOn
    End With

    TheHdw.Wait CDbl(WaitTime)
    MeasVal = TheHdw.PPMU.pins(MeasPin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    With TheHdw.PPMU.pins(MeasPin)
        .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
        TheHdw.Wait (0.3 * ms) '20191002 CT add to solve MeasI clamp for GPIO DS tests
        .Disconnect
        .Gate = tlOff
    End With

    TheHdw.Wait (0.1 * ms)
    TheHdw.Digital.pins(MeasPin).Connect

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureCurr_PPMU") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureCurr_UVI80(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceV As String, ByVal MeterCurrRange As String, Optional ByVal WaitTime As String = "0.1") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureCurr_UVI80"
    
    Dim typ_SaveCond As Harvest_ALTBackup_Type
    
    typ_SaveCond.Current = CDbl(FormatNumber(TheHdw.DCVI.pins(MeasPin).Current, 3))
    typ_SaveCond.SrcCurrentRange = CDbl(FormatNumber(TheHdw.DCVI.pins(MeasPin).CurrentRange.value, 3))
    If TheHdw.DCVI.pins(MeasPin).Gate = True Then
        typ_SaveCond.IfPowerPin = True
        typ_SaveCond.Voltage = CDbl(FormatNumber(TheHdw.DCVI.pins(MeasPin).Voltage, 3))
    End If

    With TheHdw.DCVI.pins(MeasPin)
        If TheHdw.DCVI.pins(MeasPin).Gate = False Then
            .Gate = False
            .mode = tlDCVIModeVoltage
            .Voltage = CDbl(ForceV)
        End If
        .VoltageRange.Autorange = True
        .CurrentRange.Autorange = True
        .Current = pc_Def_UVI80_Init_MeasCurrRange 'Init Current set to 2A, From Sicily, 20200423, Oscar
        .Connect tlDCVIConnectDefault
        If .Gate = False Then 'Gate On only when Gate off, From Sicily, 20200423, Oscar
            .Gate(tlDCVIGateHiZ) = False 'Added by Kaino on 20190902 for Mode alarm
            .Gate = True
        End If
    End With
        
    TheHdw.Wait 0.001
        
    With TheHdw.DCVI.pins(MeasPin)
        .Meter.mode = tlDCVIMeterCurrent
        .SetCurrentAndRange CDbl(MeterCurrRange), CDbl(MeterCurrRange)
        .CurrentRange.Autorange = True
    End With
    
    TheHdw.Wait CDbl(WaitTime)
    MeasVal = TheHdw.DCVI.pins(MeasPin).Meter.Read(tlStrobe, 10)
    
    If typ_SaveCond.IfPowerPin = True Then
        TheHdw.DCVI.pins(MeasPin).CurrentRange.value = typ_SaveCond.SrcCurrentRange
        TheHdw.DCVI.pins(MeasPin).Current = typ_SaveCond.Current
        TheHdw.DCVI.pins(MeasPin).Voltage = typ_SaveCond.Voltage
    Else
        With TheHdw.DCVI.pins(MeasPin)
            .Voltage = 0
            .Current = typ_SaveCond.SrcCurrentRange
            .Gate = False
            .Disconnect
        End With
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureCurr_UVI80") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureFreq(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, Optional ByVal Interval As String = "0.01", Optional ByVal WaitTime As String = "0.01") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureFreq"

    Dim str_PinInstrument As String
    
    str_PinInstrument = UCase(SortPinInstrument(MeasPin))
    If (str_PinInstrument = "HSD-U") Then
        Call Harvest_MeasureFreq_PPMU(MeasVal, MeasPin, Interval, WaitTime)
    Else
        TheExec.Datalog.WriteComment funcName & " did not consider this case, please check!"
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureFreq") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureFreq_PPMU(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, Optional ByVal Interval As String = "0.01", Optional ByVal WaitTime As String = "0.01") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureFreq_PPMU"
    
    Dim pld_CounterValue As New PinListData
        
    With TheHdw.Digital.pins(MeasPin).FreqCtr
        .EventSource = VOH
        .EventSlope = Positive
        .Interval = CDbl(Interval)
        .Enable = IntervalEnable
        .Clear
        TheHdw.Wait CDbl(WaitTime)
        .start
        pld_CounterValue = .Read()
    End With
    
    MeasVal = pld_CounterValue.pins(MeasPin).divide(CDbl(Interval))

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureFreq_PPMU") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureR(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceV As String, ByVal MeterCurrRange As String, Optional ByVal WaitTime As String = "0.001") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureR"

    Dim str_PinInstrument As String
    
    str_PinInstrument = UCase(SortPinInstrument(MeasPin))
    If (str_PinInstrument = "HSD-U") Then
        Call Harvest_MeasureR_PPMU(MeasVal, MeasPin, ForceV, MeterCurrRange, WaitTime)
    Else
        TheExec.Datalog.WriteComment funcName & " did not consider this case, please check!"
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureR") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureR_PPMU(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceV As String, ByVal MeterCurrRange As String, Optional ByVal WaitTime As String = "0.001") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureR"
    
    Dim Sdbl_MeasI As New SiteDouble
    
    TheHdw.Digital.pins(MeasPin).Disconnect
    
    With TheHdw.PPMU.pins(MeasPin)
        .Gate = tlOff
        .ForceI pc_Def_PPMU_InitialValue_FI
        .ForceV CDbl(ForceV), CDbl(MeterCurrRange)
        .Connect
        .Gate = tlOn
    End With
        
    TheHdw.Wait CDbl(WaitTime)
    Sdbl_MeasI = TheHdw.PPMU.pins(MeasPin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    MeasVal = Sdbl_MeasI.Invert.Multiply(CDbl(ForceV)).Abs
        
    With TheHdw.PPMU.pins(MeasPin)
        .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Disconnect
        .Gate = tlOff
    End With

    TheHdw.Digital.pins(MeasPin).Connect

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureR_PPMU") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureZ(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceV1 As String, ByVal MeterCurrRange1 As String, ByVal ForceV2 As String, ByVal MeterCurrRange2 As String, Optional ByVal WaitTime As String = "0.001") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureZ"

    Dim str_PinInstrument As String
    
    str_PinInstrument = UCase(SortPinInstrument(MeasPin))
    If (str_PinInstrument = "HSD-U") Then
        Call Harvest_MeasureZ_PPMU(MeasVal, MeasPin, ForceV1, MeterCurrRange1, ForceV2, MeterCurrRange2, WaitTime)
    Else
        TheExec.Datalog.WriteComment funcName & " did not consider this case, please check!"
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureZ") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_MeasureZ_PPMU(ByRef MeasVal As SiteDouble, ByVal MeasPin As String, ByVal ForceV1 As String, ByVal MeterCurrRange1 As String, ByVal ForceV2 As String, ByVal MeterCurrRange2 As String, Optional ByVal WaitTime As String = "0.001") As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_MeasureZ"
    
    Dim Sdbl_MeasCurr1 As New SiteDouble
    Dim Sdbl_MeasCurr2 As New SiteDouble
    Dim dbl_DiffVolt As Double
    
    TheHdw.Digital.pins(MeasPin).Disconnect
    
    With TheHdw.PPMU.pins(MeasPin)
        .Gate = tlOff
        .ForceI pc_Def_PPMU_InitialValue_FI
        .ForceV CDbl(ForceV1), CDbl(MeterCurrRange1)
        .Connect
        .Gate = tlOn
    End With
    TheHdw.Wait CDbl(WaitTime)
    Sdbl_MeasCurr1 = TheHdw.PPMU.pins(MeasPin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    With TheHdw.PPMU.pins(MeasPin)
        .ForceV CDbl(ForceV2), CDbl(MeterCurrRange2)
    End With
    TheHdw.Wait CDbl(WaitTime)
    Sdbl_MeasCurr2 = TheHdw.PPMU.pins(MeasPin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    With TheHdw.PPMU.pins(MeasPin)
        .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Disconnect
        .Gate = tlOff
    End With
    
    TheHdw.Digital.pins(MeasPin).Connect
    
    dbl_DiffVolt = CDbl(ForceV2) - CDbl(ForceV1)
    MeasVal = Sdbl_MeasCurr2.Subtract(Sdbl_MeasCurr1).Invert.Multiply(dbl_DiffVolt).Abs
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_MeasureZ_PPMU") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function BitPos2Dec(ByVal Str_Pos2Dec As String, Temp2Dec As SiteDouble) As SiteDouble

    On Error GoTo errHandler
    Dim funcName As String: funcName = "BitPos2Dec"
    ''Str_Pos2Dec : "b10010", MSB -> LSB
    ''Str_Pos2Dec : F_GFX_HARV[4:0], MSB -> LSB
    ''Outout: sVar_ReturnDec(1) = 1, sVar_ReturnDec(4) = 1; mean position 1 & 4 of bit is one and return array as output
    Dim i As Long
    
    Dim str_temp As String
    Dim Str_TempAry() As String
    
    Dim Var_Site As Variant
    
''    Dim SdB_ReturnDec As New SiteDouble
    Dim Svar_ReturnDec() As New SiteVariant

    
    If auto_isBinaryString(Str_Pos2Dec) Then
        ''input : "b00010", MSB -> LSB
        str_temp = Replace(LCase(Str_Pos2Dec), "b", "")
        Str_TempAry = Split(str_temp, "")
        ReDim Svar_ReturnDec(UBound(Str_TempAry))
        For i = 0 To UBound(Str_TempAry)
''            If Str_TempAry(i) = 1 Then
''                SdB_ReturnDec = 1
''            Else
''                SdB_ReturnDec = 0
''            End If
            If Str_TempAry(i) = 1 Then
                Svar_ReturnDec(i) = 1
            Else
                Svar_ReturnDec(i) = 0
            End If
        Next i
    Else
        'input : F_GFX_HARV[4:0], MSB -> LSB
        'return: F_GFX_HARV4,F_GFX_HARV3,F_GFX_HARV2,F_GFX_HARV1,F_GFX_HARV0
        Dim StoreHarvKeyWord As String
        Dim TempCoreIdx As Integer
        StoreHarvKeyWord = Split(Str_Pos2Dec, "[")(0)
        str_temp = Harvest_StrExpand(Str_Pos2Dec)
        Str_TempAry = Split(str_temp, ",")
        ReDim Svar_ReturnDec(UBound(Str_TempAry))
        For Each Var_Site In TheExec.sites
            For i = 0 To UBound(Str_TempAry)
                If TheExec.sites.item(Var_Site).FlagState(Str_TempAry(i)) Then
                    TempCoreIdx = CInt(Split(Str_TempAry(i), StoreHarvKeyWord)(1))
                    Svar_ReturnDec(i) = TempCoreIdx + 1
                Else
                    Svar_ReturnDec(i) = 0
                End If
''                If TheExec.sites.Item(Var_Site).FlagState(Str_TempAry(i)) Then
''                    Svar_ReturnDec = i + 1
''                Else
''                    Svar_ReturnDec = 0
''                End If
            Next i
        Next Var_Site
    End If

    For Each Var_Site In TheExec.sites
        For i = 0 To UBound(Svar_ReturnDec)
            If Svar_ReturnDec(i) > 0 Then
                Temp2Dec = Temp2Dec + Svar_ReturnDec(i)
'                BitPos2Dec = BitPos2Dec + 2 ^ (UBound(Svar_ReturnDec) - i)
            End If
        Next i
    Next Var_Site

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "BitPos2Dec") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Dec2BitPos(ByVal SDbl_Dec2Pos As SiteDouble, ByVal bitwidth As Long, sVar_ReturnPos As SiteVariant) As SiteVariant

    On Error GoTo errHandler
    Dim funcName As String: funcName = "Dec2BitPos"
    ''SDbl_Dec2Pos : 12
    ''BitWidth     : 4
    ''Output       : "1100"
    Dim i As Long
    
    Dim Var_Site As Variant
    
'    Dim sVar_ReturnPos() As Long
    
    Dim str_temp As String: str_temp = vbNullString
    
'    ReDim sVar_ReturnPos(BitWidth)
    
    For Each Var_Site In TheExec.sites
'        Dec2BitPos = auto_Dec2Bin_EFuse(SDbl_Dec2Pos(Var_Site), BitWidth, sVar_ReturnPos)
        str_temp = vbNullString
        For i = 0 To bitwidth
            If i = SDbl_Dec2Pos - 1 Then
                str_temp = str_temp & "1"
            Else
                str_temp = str_temp & "0"
            End If
        Next i
        sVar_ReturnPos = StrReverse(str_temp)
    Next Var_Site

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Dec2BitPos") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



Public Function JudgeCustomFun(Optional FailFlagOrValueArr As String, Optional CustomFuncName As String, Optional TempSdbl_TransValue As SiteDouble, Optional TempSVar_TransValue As SiteVariant) As Long
    
    On Error GoTo errHandler
    Dim funcName As String: funcName = "JudgeCustomFun"
    
    
    Dim inArr_TempFlag() As String
    Dim i, j, k As Long
    Dim DefaultStr As String
    Dim OutputStrArr() As String
    Dim TempHarvStr As String
    Dim StarIdx, EndIdx As Long
    Dim StepOrder As Long
    Dim FinalHarvStr As String: FinalHarvStr = vbNullString
    Dim StoreFlagIdx As Long: StoreFlagIdx = 0
    Dim tempSlng As New SiteLong
    Dim site As Variant
    
    TempSdbl_TransValue = 0
    
    Select Case UCase(CustomFuncName):
        
        Case "BITPOS2DEC":
            Call BitPos2Dec(FailFlagOrValueArr, TempSdbl_TransValue)
            
        Case "DEC2BITPOS"
            inArr_TempFlag = Split(FailFlagOrValueArr, ",")
            Call Dec2BitPos(TempSdbl_TransValue, UBound(inArr_TempFlag), TempSVar_TransValue)
        
        Case "CONCATENATE"
            inArr_TempFlag = Split(FailFlagOrValueArr, "+")
            For i = 0 To UBound(inArr_TempFlag)
                DefaultStr = mid(inArr_TempFlag(i), 1, InStr(inArr_TempFlag(i), "[") - 1)
                StarIdx = CLng(mid(inArr_TempFlag(i), InStr(inArr_TempFlag(i), "[") + 1, InStr(inArr_TempFlag(i), ":") - InStr(inArr_TempFlag(i), "[") - 1))
                EndIdx = CLng(mid(inArr_TempFlag(i), InStr(inArr_TempFlag(i), ":") + 1, InStr(inArr_TempFlag(i), "]") - InStr(inArr_TempFlag(i), ":") - 1))
                If StarIdx > EndIdx Then
                    StepOrder = -1
                Else
                    StepOrder = 1
                End If
                ReDim Preserve OutputStrArr(Abs(StarIdx - EndIdx + StoreFlagIdx))
                For j = StoreFlagIdx To UBound(OutputStrArr)
                    OutputStrArr(j) = DefaultStr & CStr(StarIdx + j * StepOrder + StoreFlagIdx)
                Next j
                StoreFlagIdx = UBound(OutputStrArr) + 1
            Next i

            For k = 0 To UBound(OutputStrArr)
                tempSlng = Harvest_GetAllSiteFlagState(OutputStrArr(k), 1, 0, 1)
                tempSlng = tempSlng.Multiply(Application.WorksheetFunction.power(2, UBound(OutputStrArr) - k))
                TempSdbl_TransValue = TempSdbl_TransValue.Add(tempSlng)
            Next k
        
        Case "OR"
            inArr_TempFlag = Split(FailFlagOrValueArr, "+")
            For i = 0 To UBound(inArr_TempFlag)
                tempSlng = 0
                DefaultStr = mid(inArr_TempFlag(i), 1, InStr(inArr_TempFlag(i), "[") - 1)
                StarIdx = CLng(mid(inArr_TempFlag(i), InStr(inArr_TempFlag(i), "[") + 1, InStr(inArr_TempFlag(i), ":") - InStr(inArr_TempFlag(i), "[") - 1))
                EndIdx = CLng(mid(inArr_TempFlag(i), InStr(inArr_TempFlag(i), ":") + 1, InStr(inArr_TempFlag(i), "]") - InStr(inArr_TempFlag(i), ":") - 1))
                If StarIdx > EndIdx Then
                    StepOrder = -1
                Else
                    StepOrder = 1
                End If
                ReDim Preserve OutputStrArr(Abs(StarIdx - EndIdx + StoreFlagIdx))
                For j = StoreFlagIdx To UBound(OutputStrArr)
                    OutputStrArr(j) = DefaultStr & CStr(StarIdx + j * StepOrder + StoreFlagIdx)
                    tempSlng = Harvest_GetAllSiteFlagState(OutputStrArr(j), 1, 0, -1)
                    For Each site In TheExec.sites.Active
                        If tempSlng(site) = 1 Then
                            TempSdbl_TransValue(site) = 1
                        End If
                    Next site
                Next j
                StoreFlagIdx = UBound(OutputStrArr) + 1
            Next i
           
            
    End Select

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "JudgeCustomFun") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
Public Function Printing_StandalonePat(pat As String, patset() As String, Optional patcnt As Long = 0)
    On Error GoTo errHandler
    Dim funcName As String: funcName = "Printing_StandalonePat"
    Dim site As Variant
    Dim DigitalPin() As String
    Dim TempPat() As String
    Dim i, j As Integer
    Dim FailCount As Long: FailCount = 0
    Dim Count As Long
    Dim FFC, failNum As Long
    Dim maxDepth As Integer
    Dim printFailInfor As Boolean: printFailInfor = True

    If printFailInfor = True Then
        TempPat = Split(pat, ":")
        For Each site In TheExec.sites.Active
            FailCount = 0
            Count = TheHdw.Digital.FailedPinsCount(site)
            If Count <> 0 Then
                FFC = TheHdw.Digital.hram.PatGenInfo(0, pgCycle)
                DigitalPin = TheHdw.Digital.FailedPins(site)
                For i = 0 To UBound(DigitalPin)
                    If patcnt = 0 Then
                        TheExec.Datalog.WriteComment "site " & site & "  PatternName:" & TempPat(1) & ", FailingPin:" & DigitalPin(i)
                    Else
                        TheExec.Datalog.WriteComment "site " & site & "  PatternName:" & pat & ", FailingPin:" & DigitalPin(i)
                    End If
                    FailCount = TheHdw.Digital.pins(DigitalPin(i)).FailCount + FailCount
                Next i
                If patcnt = 0 Then
                    TheExec.Datalog.WriteComment "site " & site & "  PatternName:" & TempPat(1) & ", Total Failed Cycle : " & FailCount & ", First Failed Cycle : " & FFC
                Else
                    TheExec.Datalog.WriteComment "site " & site & "  PatternName:" & pat & ", Total Failed Cycle : " & FailCount & ", First Failed Cycle : " & FFC
                End If
            End If
        Next site
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Printing_StandalonePat") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



Public Function CheckInitPat(PatName As String, TempCheckInitPat As Boolean)


    On Error GoTo errHandler
    Dim funcName As String: funcName = "CheckInitPat"

    Dim Patinfor() As String
    Dim TempPatInfor() As String

    Patinfor = Split(LCase(PatName), "\")
    TempPatInfor = Split(LCase(Patinfor(UBound(Patinfor))), "_")
    TempCheckInitPat = False
    
    If TempPatInfor(3) Like "*in*" Then
        TempCheckInitPat = True
    End If
        
             
    ' CJR: 4/23/24: Check if PL pattern matches pattern keyword in HarvestPingFlag_Table
'    Dim i As Long
'    If LCase(TempPatInfor(3)) Like "*pl*" Then
'        For i = 0 To UBound(HarvPinFlagMapping)
'            If UCase(PatName) Like UCase(HarvPinFlagMapping(i).Pattern) Then    '' ex: "*CCC0*" matched
'                TempCheckInitPat = False
'                Exit For
'            Else
'                TempCheckInitPat = True
'            End If
'        Next i
'    End If
    ' End CJR:4/23/24 Updated
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "CheckInitPat") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
        

End Function



Public Function HarvCustomFeature(ConditionArr() As String, failFlagArr() As String, CustomCondition() As Boolean) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "HarvCustomFeature"
    
    
    Dim i As Integer
    Dim site As Variant
    Dim TempPinGupStr() As String
    
    For i = 0 To UBound(ConditionArr)
        If CustomCondition(i) = True Then
            TempPinGupStr = Split(ConditionArr(i), ":")
            For Each site In TheExec.sites
                If TheExec.sites.item(site).FlagState(failFlagArr(i)) = logicTrue Then
                    TheHdw.Digital.pins(TempPinGupStr(1)).DisableCompare = True
                    TheExec.Datalog.WriteComment "--(Enable Pin mask Feature)--" & "site = " & site
                End If
            Next site
        End If
    Next i
    

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "HarvCustomFeature") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function initHarvSumFlag(Harv_FailCoreSumFlag As String, Harv_AllCorePassFlag As String, Harv_BinOutFailFlag As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "initHarvSumFlag"

    Dim TempFailCoreSum() As String
    Dim tempAllCorePass() As String
    Dim tempBinOutFail() As String
    Dim i As Long
    Dim TempAryFailCore() As String
    Dim j As Long
    Dim site As Variant

    TempFailCoreSum = Split(Harv_FailCoreSumFlag, "|")
    tempAllCorePass = Split(Harv_AllCorePassFlag, "|")
    tempBinOutFail = Split(Harv_BinOutFailFlag, "|")
    
    For Each site In TheExec.sites.Active
        For i = 0 To UBound(TempFailCoreSum)
            If InStr(1, TempFailCoreSum(i), "+") <> 0 Then
                TempAryFailCore = Split(TempFailCoreSum(i), "+")
                For j = 0 To UBound(TempAryFailCore)
                    TheExec.sites.item(site).FlagState(TempAryFailCore(j)) = logicFalse
                Next j
            Else
                TheExec.sites.item(site).FlagState(TempFailCoreSum(i)) = logicFalse
                TheExec.sites.item(site).FlagState(tempAllCorePass(i)) = logicFalse
                TheExec.sites.item(site).FlagState(tempBinOutFail(i)) = logicFalse
            End If
        Next i
    Next site

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "initHarvSumFlag") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



Public Function Harvest_CMEM_PostResult_UP(ByRef Failflag As String, ByVal CheckPins As String, Optional ByVal CheckVectors As String = vbNullString, Optional PatName As String, Optional CustomCondition As Boolean) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_CMEM_PostResult_UP"
    
    Dim sInstName As String
    Dim vsite As Variant
    
    Dim str_CheckPinsArr() As String
    Dim lng_CheckPinsNum As Long
    Dim str_CheckPins_ReUnite As String
    
    Dim Sbln_PatternPass As New SiteBoolean
    Dim SiteIndexData  As New SiteVariant
    Dim SitePinData  As New SiteVariant
    Dim patternNames() As String
    Dim lastFailPerPattern() As Long
'''''    Dim dbl_CmemCycleDataArr() As Double 'Data type is Double because cycle count is 40 bits.
'''''    Dim dbl_CmemVectorDataArr() As Double
    Dim dbl_CmemCycleDataArr As New SiteVariant
    Dim dbl_CmemVectorDataArr As New SiteVariant
    
    Dim indexArr() As Long 'Used in site loop
    Dim pinDataArr() As Double 'Used in site loop
    Dim bigPinList As Integer 'Used when pinlist is greater than 32.
    Dim patIdx As Long
    Dim pinStr As String
    Dim PinDict As New Dictionary
    Dim Slng_FailFlag As New SiteLong
    Dim str_CheckVectorsArr() As String
    Dim dbl_TempFailVector As Double
    Dim lng_FailDetailStartIndex As Long
    Dim i As Long
    Dim k As Long
    Dim j As Integer
    Dim dbl_MinCmemIndex As Double
    Dim Temp_CheckPinsArr() As String
    Dim Temp_str_CheckPinsArr() As String
    
    Dim CheckHarvPinGupBool As New SiteBoolean: CheckHarvPinGupBool = False
    Dim CntHarvFailCycle As New SiteDouble: CntHarvFailCycle = 0
    
    Dim TempTnameConver As String
    
    Dim FailCount As New PinListData
        
    sInstName = TheExec.DataManager.instancename
    
    Call TheExec.DataManager.DecomposePinList(CheckPins, str_CheckPinsArr, lng_CheckPinsNum)
    
    If UBound(str_CheckPinsArr) < 31 Then
        ReDim Temp_CheckPinsArr(0)
        Temp_CheckPinsArr(0) = Join(str_CheckPinsArr, ",")
    Else
        bigPinList = Ceiling(CDbl((UBound(str_CheckPinsArr) + 1) / 32))
        ReDim Temp_CheckPinsArr(bigPinList - 1)
        For k = 0 To bigPinList - 1
            For i = 0 To 31
                If i + k * 32 > UBound(str_CheckPinsArr) Then Exit For
                If i = 0 Then
                    Temp_CheckPinsArr(k) = str_CheckPinsArr(i + 32 * k)
                Else
                    Temp_CheckPinsArr(k) = Temp_CheckPinsArr(k) + "," + str_CheckPinsArr(i + 32 * k)
                End If
            Next i
        Next k
    End If
    
    Sbln_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    If Sbln_PatternPass.Any(False) Then
        'Read failing pins for all sites, compressing the failures.
'''''        Call TheHdw.Digital.Pins(str_CheckPins_ReUnite).CMEM.StoredCycleData(SiteIndexData, SitePinData, -1, True)
        'Read back the last index where data was captured for each pattern.
'''''        Call TheHdw.Digital.CMEM.PatternName(lastFailPerPattern, patternNames)
        'Read back the central data.
'''''        dbl_CmemCycleDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMModCycle, -1)
'''''        dbl_CmemVectorDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMVMVectorOffset, -1)
        
        For Each vsite In TheExec.sites.Selected
            If Sbln_PatternPass(vsite) = False Then
                Call TheHdw.Digital.CMEM.PatternName(lastFailPerPattern, patternNames)
                'Read back the central data.
                dbl_CmemCycleDataArr(vsite) = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMModCycle, -1)
                dbl_CmemVectorDataArr(vsite) = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMVMVectorOffset, -1)
                dbl_MinCmemIndex = Min(UBound(dbl_CmemCycleDataArr(vsite)), UBound(dbl_CmemVectorDataArr(vsite)))
        ''        Read failing pins for all sites, compressing the failures.'
                For j = 0 To UBound(Temp_CheckPinsArr)
                    Temp_str_CheckPinsArr = Split(Temp_CheckPinsArr(j), ",")
                    Call TheHdw.Digital.pins(Temp_CheckPinsArr(j)).CMEM.StoredCycleData(SiteIndexData, SitePinData, -1, True)
                
                    'Move siteVariant to normal array of Long/Double
                    indexArr = SiteIndexData
                    pinDataArr = SitePinData
                    'pat fail at pinA, pinB. but capture pinC => UBound(IndexArr) = -1
                    If UBound(indexArr) <> -1 Then
                        'being here means capture pins did fail
                        If CheckVectors = "" Then
'                            TheExec.Datalog.WriteComment "Site" & vSite & " Pin : " & CheckPins & " total fail cycle = " & CStr(UBound(indexArr) + 1) & "."
                            TheExec.sites.item(vsite).FlagState(Failflag) = logicTrue
                        Else
                            str_CheckVectorsArr = Split(CheckVectors, ",")
                            dbl_TempFailVector = -1
                            For i = LBound(indexArr) To UBound(indexArr)
                                If dbl_TempFailVector <> dbl_CmemVectorDataArr(indexArr(i)) Then
                                    dbl_TempFailVector = dbl_CmemVectorDataArr(indexArr(i))
                                    For k = 0 To UBound(str_CheckVectorsArr)
                                        If dbl_TempFailVector = str_CheckVectorsArr(k) Then
                                            TheExec.Datalog.WriteComment "Site" & vsite & " Pin : " & CheckPins & " fail at vector" & CStr(dbl_TempFailVector) & "."
                                            TheExec.sites.item(vsite).FlagState(Failflag) = logicTrue
                                            Exit For
                                        Else
                                        End If
                                    Next k
                                Else
                                End If
                            Next i
                        End If
                        
                        'the following is for fail pin and fail vector parsing
                        Const ShowFailDetail = True
                        Const ShowFailDetail_MaxNum = 511 '-1 for all
                        If ShowFailDetail Then
                            'bigPinlist will equal 1 for <32 pins; 2 for 64 to 33 pins; and so on.
    '                        bigPinList = (UBound(pinDataArr) + 1) \ (UBound(indexArr) + 1)
                            patIdx = 0
                            If ShowFailDetail_MaxNum > -1 Then
                                If j = 0 Then TheExec.Datalog.WriteComment "Print out all fail cycle."
                                lng_FailDetailStartIndex = LBound(indexArr)
                            ElseIf UBound(indexArr) > ShowFailDetail_MaxNum Then
                                If j = 0 Then TheExec.Datalog.WriteComment "Fail Cycle over " & ShowFailDetail_MaxNum & ", Only print the last " & ShowFailDetail_MaxNum & "."
                                lng_FailDetailStartIndex = UBound(indexArr) - ShowFailDetail_MaxNum
                            Else
                                lng_FailDetailStartIndex = LBound(indexArr)
                            End If
                            'For k = 0 To bigPinList - 1
                                'If bigPinList > 1 Then lng_FailDetailStartIndex = UBound(indexArr)
                            For i = lng_FailDetailStartIndex To UBound(indexArr)
                                pinStr = vbNullString
                                PinDict.RemoveAll
    '                            For k = 0 To bigPinList - 1
    '                            PinStr = FailingPins(pinDataArr(i * bigPinList + k), str_CheckPinsArr, k, PinDict) & "," & PinStr
                                pinStr = FailingPins(pinDataArr(i), Temp_str_CheckPinsArr, 0, PinDict) & "," & pinStr
    '                            Next k
                                If patIdx > UBound(patternNames) Then
                                    'Raise error. This should never happen.
                                End If
                                If i > lastFailPerPattern(patIdx) Then
                                    'Go to the next pattern in the list.
                                    patIdx = patIdx + 1
                                End If
                                
                                If i < ShowFailDetail_MaxNum Then 'And pinDataArr(i) <> 0 Then
                                    TheExec.Datalog.WriteComment "vector : " & CStr(dbl_CmemVectorDataArr(vsite)(indexArr(i))) + _
                                                                 " cycle : " & CStr(dbl_CmemCycleDataArr(vsite)(indexArr(i))) + _
                                                                 " pin : " & pinStr
                                End If
                            Next i
                            'Next k
''''                            CheckHarvPinGupBool(vsite) = True
''''                            CntHarvFailCycle(vsite) = CntHarvFailCycle(vsite) + (UBound(indexArr) + 1)
                        End If
                        CheckHarvPinGupBool(vsite) = True
                        CntHarvFailCycle(vsite) = CntHarvFailCycle(vsite) + (UBound(indexArr) + 1)
'''''''                        Dim FailCount As New PinListData ''''' nop for binout NonHarv Pin fail 20230216
'''''''                        FailCount = TheHdw.Digital.Pins(Temp_CheckPinsArr(j)).FailCount
'''''''                        For i = 0 To FailCount.Pins.Count - 1
'''''''                            If FailCount.Pins.item(i).value <> 0 Then
'''''''                                HarvOtherFailCnt(vsite) = HarvOtherFailCnt(vsite) + 1
'''''''                            End If
'''''''                        Next i
                    Else
                        If CheckHarvPinGupBool(vsite) = True Then
                        Else
                            CheckHarvPinGupBool(vsite) = False
                        End If
                        'CMEM capture pins didnot fail
''''                        TheExec.Datalog.WriteComment "Site" & vSite & " Pat Fail, but Pin : " & CheckPins & " did not fail."
''''                        If TheExec.sites.item(vSite).FlagState(Failflag) = logicClear Then
''''                            TheExec.sites.item(vSite).FlagState(Failflag) = logicFalse
''''                        ElseIf TheExec.sites.item(vSite).FlagState(Failflag) = logicTrue Then
''''                            'need check?
''''                        End If
                    End If
                    
                    FailCount = TheHdw.Digital.pins(Temp_CheckPinsArr(j)).FailCount ''''' Counter the harvesting pin fail to do the comparison
                    For i = 0 To FailCount.pins.Count - 1
                        If FailCount.pins.item(i).value <> 0 Then
                            HarvPinsFailCnt(vsite) = HarvPinsFailCnt(vsite) + 1
                        End If
                    Next i
                    
                Next j
                    
                If CheckHarvPinGupBool(vsite) = False Then
                    If CustomCondition = True Then
                        TheExec.Datalog.WriteComment "Site" & vsite & " Pat Fail, and Pin : " & CheckPins & " set DisableCompare."
                    Else
                        TheExec.Datalog.WriteComment "Site" & vsite & " Pat Fail, but Pin : " & CheckPins & " did not fail."
                    End If
                    If TheExec.sites.item(vsite).FlagState(Failflag) = logicClear Then
                        TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse
                    ElseIf TheExec.sites.item(vsite).FlagState(Failflag) = logicTrue Then
                        'need check?
                    ElseIf TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse Then
                        ''''''HarvFailCnt(vsite) = HarvFailCnt(vsite) - 1 ''''' nop for binout NonHarv Pin fail 20230216
                    End If
                Else
                    TheExec.Datalog.WriteComment "Site" & vsite & " Pin : " & CheckPins & " total fail cycle = " & CStr(CntHarvFailCycle(vsite)) & "."
                End If
            End If

            Slng_FailFlag = TheExec.sites.item(vsite).FlagState(Failflag)
        Next vsite
    Else
        'all site pass
        'theexec.Datalog.WriteComment "All Site Pat Pass."
        For Each vsite In TheExec.sites.Selected
            If TheExec.sites.item(vsite).FlagState(Failflag) = logicClear Then
                TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse
            End If
            Slng_FailFlag = TheExec.sites.item(vsite).FlagState(Failflag)
        Next vsite
    End If
        
    TheExec.flow.TestLimit Slng_FailFlag, 0, 1, Tname:=sInstName & "_" & CheckPins
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Harvest", "Harvest_CMEM_PostResult_UP") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_CMEM_PostResult_UF(ByRef Failflag As String, ByVal CheckPins As String, Optional ByVal CheckVectors As String = vbNullString) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_CMEM_PostResult_UF"
    
    Dim sInstName As String
    Dim vsite As Variant
    
    Dim str_CheckPinsArr() As String
    Dim lng_CheckPinsNum As Long
    Dim str_CheckPins_ReUnite As String
    
    Dim Sbln_PatternPass As New SiteBoolean
    Dim SiteIndexData  As New SiteVariant
    Dim SitePinData  As New SiteVariant
    Dim patternNames() As String
    Dim lastFailPerPattern() As Long
    Dim dbl_CmemCycleDataArr() As Double 'Data type is Double because cycle count is 40 bits.
    Dim dbl_CmemVectorDataArr() As Double
    Dim indexArr() As Long 'Used in site loop
    Dim pinDataArr() As Double 'Used in site loop
    Dim bigPinList As Integer 'Used when pinlist is greater than 32.
    Dim patIdx As Long
    Dim pinStr As String
    Dim PinDict As New Dictionary
    Dim Slng_FailFlag As New SiteLong
    Dim str_CheckVectorsArr() As String
    Dim dbl_TempFailVector As Double
    Dim lng_FailDetailStartIndex As Long
    Dim i As Long
    Dim k As Long
    Dim dbl_MinCmemIndex As Double
    Dim FailCount As New PinListData
    Dim HarvPinsFailCntCheck As New SiteLong 'Check Harvest Flag Status, if HarvPinsFailCntCheck >0 & Harvest Flag Status is False then Set HarvPins Fail Flag to TRUE
    
    sInstName = TheExec.DataManager.instancename
    
    Call TheExec.DataManager.DecomposePinList(CheckPins, str_CheckPinsArr, lng_CheckPinsNum)
    str_CheckPins_ReUnite = Join(str_CheckPinsArr, ",")
    
    Sbln_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    If Sbln_PatternPass.Any(False) Then
        'Read failing pins for all sites, compressing the failures.
'        Call TheHdw.Digital.Pins(str_CheckPins_ReUnite).CMEM.StoredCycleData(SiteIndexData, SitePinData, -1, True)
        'Read back the last index where data was captured for each pattern.
        Call TheHdw.Digital.CMEM.PatternName(lastFailPerPattern, patternNames)
        'Read back the central data.
        dbl_CmemCycleDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMModCycle, -1)
        dbl_CmemVectorDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMVMVectorOffset, -1)
        dbl_MinCmemIndex = Min(UBound(dbl_CmemCycleDataArr), UBound(dbl_CmemVectorDataArr))
''        Read failing pins for all sites, compressing the failures.
        Call TheHdw.Digital.pins(str_CheckPins_ReUnite).CMEM.StoredCycleData(SiteIndexData, SitePinData, dbl_MinCmemIndex + 1, True)
   
        
        For Each vsite In TheExec.sites.Selected
            If Sbln_PatternPass(vsite) = False Then
                'Move siteVariant to normal array of Long/Double
                indexArr = SiteIndexData
                pinDataArr = SitePinData
                'pat fail at pinA, pinB. but capture pinC => UBound(IndexArr) = -1
                If UBound(indexArr) <> -1 Then
                    'being here means capture pins did fail
                    If CheckVectors = "" Then
                        TheExec.Datalog.WriteComment "Site" & vsite & " Pin : " & CheckPins & " total fail cycle = " & CStr(UBound(indexArr) + 1) & "."
                        TheExec.sites.item(vsite).FlagState(Failflag) = logicTrue
                    Else
                        str_CheckVectorsArr = Split(CheckVectors, ",")
                        dbl_TempFailVector = -1
                        For i = LBound(indexArr) To UBound(indexArr)
                            If dbl_TempFailVector <> dbl_CmemVectorDataArr(indexArr(i)) Then
                                dbl_TempFailVector = dbl_CmemVectorDataArr(indexArr(i))
                                For k = 0 To UBound(str_CheckVectorsArr)
                                    If dbl_TempFailVector = str_CheckVectorsArr(k) Then
                                        TheExec.Datalog.WriteComment "Site" & vsite & " Pin : " & CheckPins & " fail at vector" & CStr(dbl_TempFailVector) & "."
                                        TheExec.sites.item(vsite).FlagState(Failflag) = logicTrue
                                        Exit For
                                    Else
                                    End If
                                Next k
                            Else
                            End If
                        Next i
                    End If

                    'the following is for fail pin and fail vector parsing
                    Const ShowFailDetail = True ' true enable fail cycle print
'                    Const ShowFailDetail_MaxNum = 1 '-1 for all
                    Const ShowFailDetail_MaxNum = 511 ' Setting same as Ellis
                    If ShowFailDetail Then
                        'bigPinlist will equal 1 for <32 pins; 2 for 64 to 33 pins; and so on.
                        bigPinList = (UBound(pinDataArr) + 1) \ (UBound(indexArr) + 1)
                        patIdx = 0
                        If ShowFailDetail_MaxNum > -1 Then
                            TheExec.Datalog.WriteComment "Print out all fail cycle."
                            lng_FailDetailStartIndex = LBound(indexArr)
                        ElseIf UBound(indexArr) > ShowFailDetail_MaxNum Then
                            TheExec.Datalog.WriteComment "Fail Cycle over " & ShowFailDetail_MaxNum & ", Only print the last " & ShowFailDetail_MaxNum & "."
                            lng_FailDetailStartIndex = UBound(indexArr) - ShowFailDetail_MaxNum
                        Else
                            lng_FailDetailStartIndex = LBound(indexArr)
                        End If
''                        Dim l_temp As Long
''                        If UBound(indexArr) > ShowFailDetail_MaxNum Then
''                            l_temp = ShowFailDetail_MaxNum
''                        Else
''                            l_temp = UBound(indexArr)
''                        End If
                        For i = lng_FailDetailStartIndex To UBound(indexArr)
                            pinStr = vbNullString
                            PinDict.RemoveAll
                            For k = 0 To bigPinList - 1
                                pinStr = FailingPins(pinDataArr(i * bigPinList + k), str_CheckPinsArr, k, PinDict) & "," & pinStr
                            Next k
                            If patIdx > UBound(patternNames) Then
                                'Raise error. This should never happen.
                            End If
                            If i > lastFailPerPattern(patIdx) Then
                                'Go to the next pattern in the list.
                                patIdx = patIdx + 1
                            End If

                            If i >= ShowFailDetail_MaxNum Then Exit For

                            If i < ShowFailDetail_MaxNum Then
                                TheExec.Datalog.WriteComment "vector : " & CStr(dbl_CmemVectorDataArr(indexArr(i))) + _
                                                             " cycle : " & CStr(dbl_CmemCycleDataArr(indexArr(i))) + _
                                                             " pin : " & pinStr
                            End If
                        Next i
                    End If
                Else
                    'CMEM capture pins didnot fail
                    TheExec.Datalog.WriteComment "Site" & vsite & " Pat Fail, but Pin : " & CheckPins & " did not fail."
                    If TheExec.sites.item(vsite).FlagState(Failflag) = logicClear Then
                        TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse
                    ElseIf TheExec.sites.item(vsite).FlagState(Failflag) = logicTrue Then
                        'need check?
                    ElseIf TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse Then
                        ''''''HarvFailCnt(vsite) = HarvFailCnt(vsite) - 1
                    End If
                End If
                
            End If
            If TheExec.sites.item(vsite).FlagState(Failflag) = logicClear Then
                TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse
            ElseIf TheExec.sites.item(vsite).FlagState(Failflag) = logicTrue Then
                'need check?
            ElseIf TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse Then
                ''''''HarvFailCnt(vsite) = HarvFailCnt(vsite) - 1
            End If
            
            Slng_FailFlag = TheExec.sites.item(vsite).FlagState(Failflag)
        Next vsite
    Else
        'all site pass
        'theexec.Datalog.WriteComment "All Site Pat Pass."
        For Each vsite In TheExec.sites.Selected
            If TheExec.sites.item(vsite).FlagState(Failflag) = logicClear Then
                TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse
            End If
            Slng_FailFlag = TheExec.sites.item(vsite).FlagState(Failflag)
        Next vsite
    End If
        
    
    
    For Each vsite In TheExec.sites.Selected
        FailCount = TheHdw.Digital.pins(str_CheckPins_ReUnite).FailCount
        For i = 0 To FailCount.pins.Count - 1
            If FailCount.pins.item(i).value <> 0 Then
                HarvPinsFailCnt(vsite) = HarvPinsFailCnt(vsite) + 1
                HarvPinsFailCntCheck(vsite) = HarvPinsFailCntCheck(vsite) + 1
            End If
        Next i
        If HarvPinsFailCntCheck(vsite) > 0 Then
               If TheExec.sites.item(vsite).FlagState(Failflag) = logicFalse Then
                  TheExec.sites.item(vsite).FlagState(Failflag) = logicTrue
                  Slng_FailFlag = TheExec.sites.item(vsite).FlagState(Failflag) 'Keep this code above VBT code logic, VBT code only reset "Slng_FailFlag" for changed Harvest Fail Site Flag (Only add this line code)
                      TheExec.Datalog.WriteComment "Site" & vsite & " Harvest Pin : " & CheckPins & " Fail, But Harvest Fail Flag: " & Failflag & " is False, Set Harvest Fail Flag: " & Failflag & " to TRUE"
               End If
        End If
    Next vsite
    HarvPinsFailCntCheck = 0
    TheExec.flow.TestLimit Slng_FailFlag, 0, 1, Tname:=sInstName & "_" & CheckPins
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Parsing_ATE_STR_Summary_Table()
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim i As Long, j As Long, k As Long
    Dim Row_of_Title As Integer
    Dim col_PTR_TestName As Integer
    Dim col_Operator As Integer
    Dim col_InputValue As Integer
    Dim col_harvotherpin As Integer
    Dim blockCnt As Integer

    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
    On Error GoTo errHandler

    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "ATE_STR_Summary"
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        enableRowParsing = False
        
        '''//Check the header of the table
        '''Get the columns for the diverse coefficient
        For i = 1 To MaxRow
            For j = 1 To maxcol
                If UCase(ws_def.Cells(i, j).value) = UCase("PTR Test Name") Then
                    col_PTR_TestName = j
                    Row_of_Title = i
                End If
                
                If Row_of_Title > 0 Then
                    If UCase(ws_def.Cells(Row_of_Title, j).value) = UCase("Operator") Then
                        col_Operator = j
                    ElseIf UCase(ws_def.Cells(Row_of_Title, j).value) = UCase("Value") Then
                        col_InputValue = j
                    Else
                        '''Do nothing
                    End If
                End If
            Next j
            
            If Row_of_Title > 0 Then
                If col_PTR_TestName > 0 And col_Operator > 0 And col_InputValue > 0 Then
                    enableRowParsing = True
                Else
                    enableRowParsing = False
                    
                    If col_PTR_TestName = 0 Then
                        TheExec.Datalog.WriteComment "Column PTR Test Name doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column PTR Test Name doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_Operator = 0 Then
                        TheExec.Datalog.WriteComment "Column Operator doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column Operator doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_InputValue = 0 Then
                        TheExec.Datalog.WriteComment "Column Value doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column Value doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                End If
                Exit For
            End If
        Next i
        
        blockCnt = 0
        
        '''//Start parsing the cells
        If enableRowParsing = True Then
            For i = Row_of_Title + 1 To MaxRow
                '''//Check if the row is "End" or empty.
                If LCase(ws_def.Cells(i, 1).value) = "end" Then
                    Exit For
                End If
                
                If LCase(ws_def.Cells(i, col_PTR_TestName).value) <> "" Then
                    ReDim Preserve PTR_Flag(blockCnt)
    '                PTR_Flag(blockCnt).Enable_otherPinGrp = False
                    
                    PTR_Flag(blockCnt).PTR_Test_Name = ws_def.Cells(i, col_PTR_TestName).value
                    PTR_Flag(blockCnt).operator = ws_def.Cells(i, col_Operator).value
                    PTR_Flag(blockCnt).value = ws_def.Cells(i, col_InputValue).value
    '                If col_harvotherpin > 0 Then        '' other pin grp header is optional
    '                    HarvPinFlagMapping(blockCnt).HarvPinGrpComFailArr = Split(ws_def.Cells(i, col_harvotherpin).Value, ";")
    '                    HarvPinFlagMapping(blockCnt).Enable_otherPinGrp = True
    '                End If
                    
                    blockCnt = blockCnt + 1
                End If
                ''20220711: Added for optional use of HarvPinGrp for BinCut
'                Flag_HarvPinFlag_Mapping_Table_Parsed = True
            Next i
                        
             ATE_STR_Summary_Table_Parse_Flag = True
        Else
            TheExec.Datalog.WriteComment sheetName & " doesn't have the correct header for HarvestPinFlag_Table. Error!!!"
            TheExec.ErrorLogMessage sheetName & " doesn't have the correct header for HarvestPinFlag_Table. Error!!!"
        End If
    Else    ''If isSheetFound = False
'        Flag_HarvPinFlag_Mapping_Table_Parsed = False       ''20220711: Added for checking existance of HarvPinGrp Table
        TheExec.Datalog.WriteComment sheetName & "doesn't exist, please check it"
    End If
    
    
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_ATE_STR_Summary_Table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_ATE_STR_Summary_Table"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ATE_STR_Summary_Flag_Operate()
On Error GoTo errHandler
Dim i As Long, j As Long
Dim vsite As Variant
Dim Input_Flag_Arr() As String
Dim Flag_len As Integer
Dim Equal_pos As Integer


For i = 0 To UBound(PTR_Flag)
    
    Select Case UCase(PTR_Flag(i).operator)
        Case UCase("EQUAL")
            For Each vsite In TheExec.sites
'                TheExec.sites.item(Vsite).FlagState(PTR_Flag(i).Value) = logicTrue
                TheExec.sites.item(vsite).FlagState(Trim(PTR_Flag(i).PTR_Test_Name)) = TheExec.sites.item(vsite).FlagState(Trim(PTR_Flag(i).value))
            Next vsite
        Case UCase("OR")
            PTR_Flag(i).value = Replace(PTR_Flag(i).value, """", "")
'            PTR_Flag(i).Value = Replace(PTR_Flag(i).Value, ")", "")
            Input_Flag_Arr = Split(PTR_Flag(i).value, ",")
            
                For Each vsite In TheExec.sites.Selected
'                     If TheExec.sites.item(vsite).FlagState(Trim(PTR_Flag(i).PTR_Test_Name)) = logicClear Then
'                        TheExec.sites.item(vsite).FlagState(Trim(PTR_Flag(i).PTR_Test_Name)) = logicFalse
'                     End If
                        
                        For j = 0 To UBound(Input_Flag_Arr)
                            If TheExec.sites.item(vsite).FlagState(Trim(Input_Flag_Arr(j))) = logicTrue Then
                                TheExec.sites.item(vsite).FlagState(Trim(PTR_Flag(i).PTR_Test_Name)) = logicTrue
                                Exit For
                            End If
                        Next j
                     
                Next vsite
            
        Case UCase("IF")
           PTR_Flag(i).value = Replace(PTR_Flag(i).value, """", "")
           PTR_Flag(i).value = Replace(PTR_Flag(i).value, "(", "") 'Remove "(" & ")"
           PTR_Flag(i).value = Replace(PTR_Flag(i).value, ")", "")
           Input_Flag_Arr = Split(PTR_Flag(i).value, ",")
           Equal_pos = InStr(Input_Flag_Arr(0), "=")
           Input_Flag_Arr(0) = left(Input_Flag_Arr(0), Equal_pos - 1)
             For Each vsite In TheExec.sites.Selected
                    If TheExec.sites.item(vsite).FlagState(Trim(Input_Flag_Arr(0))) = logicFalse Then
                        TheExec.sites.item(vsite).FlagState(Trim(PTR_Flag(i).PTR_Test_Name)) = TheExec.sites.item(vsite).FlagState(Trim(Input_Flag_Arr(1)))
                    Else
                        TheExec.sites.item(vsite).FlagState(Trim(PTR_Flag(i).PTR_Test_Name)) = Trim(Input_Flag_Arr(2))
                    End If

             Next vsite
        Case Else
            
    End Select
    
'    For Each Vsite In TheExec.sites.Selected
'
'    Next Vsite
Next i


Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of ATE_STR_Summary_Flag_Operate"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of ATE_STR_Summary_Flag_Operate"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Enable_PinMaskFeature(ConditionArr() As String, failFlagArr() As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Enable_PinMaskFeature"
    
    Dim tmp_int As Integer
    Dim j As Integer
    Dim i As Integer
    Dim site As Variant
    Dim TempPinGupStr() As String
    Dim FlagCnt As Integer ' Harvest Fail flag count for all site
    Dim FlagCnt_add_Flag As Boolean 'Use this Flag to decide whether FlagCnt +1
    FlagCnt = 0 ' Initial Harvest Fail flag count
    FlagCnt_add_Flag = False ' Initial FlagCnt_add_Flag
    

    For i = 0 To UBound(ConditionArr)

        For Each site In TheExec.sites
            If TheExec.sites.item(site).FlagState(failFlagArr(i)) = logicTrue Then
               TheHdw.Digital.pins(ConditionArr(i)).DisableCompare = True
               TheExec.Datalog.WriteComment "--(Enable Pin mask Feature)-- " & "site = " & site & ", Harvest Fail Flag Name: " & failFlagArr(i) & ", Pin Name: " & ConditionArr(i)
            End If
        Next site

    Next i
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function

Public Function Enable_DisableComparePinFromFuse(ConditionArr() As String, failFlagArr() As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Enable_DisableComparePinFromFuse"
    
    Dim tmp_int As Integer
    Dim j As Integer
    Dim i As Integer
    Dim site As Variant
    Dim TempPinGupStr() As String
    Dim FlagCnt As Integer ' Harvest Fail flag count for all site
    Dim FlagCnt_add_Flag As Boolean 'Use this Flag to decide whether FlagCnt +1
    FlagCnt = 0 ' Initial Harvest Fail flag count
    FlagCnt_add_Flag = False ' Initial FlagCnt_add_Flag
    

    For i = 0 To UBound(ConditionArr)

        For Each site In TheExec.sites
            If TheExec.sites.item(site).FlagState(failFlagArr(i)) = logicTrue Then
'                If EnableCoreMask = True Then
'                    For j = 0 To UBound(Harvest_failed_FailFlag_From_Harvest_eFuse_Read)
'                        If Harvest_failed_FailFlag_From_Harvest_eFuse_Read(j)(site) = FailFlagArr(i) Then
                            If glb_Harvest_Fail_Flag_From_eFuseRead_Dict.Exists((site & failFlagArr(i))) Then
                                TheHdw.Digital.pins(ConditionArr(i)).DisableCompare = True
                                TheExec.Datalog.WriteComment "--(Enable_DisableComparePinFromFuse)-- Harvest Fail at previous stage, " & "site = " & site & ", Harvest Fail Flag Name: " & failFlagArr(i) & ", Pin Name: " & ConditionArr(i)
                            End If
'                            Exit For
'                        End If
'                    Next j
'                End If
            End If
        Next site

    Next i
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function

' This function turns off central capture at the End of Body.
' Public Function EndOfBodyIPF(argc As Integer, argv() As String)
Public Function Disable_PinMaskFeature(ConditionArr() As String, failFlagArr() As String, Optional initPatBool As Boolean) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Disable_PinMaskFeature"
    
    Dim i As Integer
    Dim TempPinGupStr() As String
    Dim site As Variant
    
    If TheExec.TesterMode = testModeOnline Then
        ' Use this statement to turn off the Pin CMEM capture.
        Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone) ' Resets CMEM
        ' Use this statement to turn off the Central CMEM capture.
        TheHdw.Digital.CMEM.CentralFields = tlCMEMNone
    End If
     
    If initPatBool = False Then
        For i = 0 To UBound(ConditionArr)
'            TempPinGupStr = Split(ConditionArr(i), ":")
'            If CustomCondition(i) = True Then
                For Each site In TheExec.sites
                    If TheExec.sites.item(site).FlagState(failFlagArr(i)) = logicTrue Then
                        TheHdw.Digital.pins(ConditionArr(i)).DisableCompare = False
                        TheExec.Datalog.WriteComment "--(Disable Pin mask Feature)--" & "site = " & site & ", Pin Name: " & ConditionArr(i)
                    End If
                Next site
'            End If
        Next i
    End If
    
     ''''' New request for pattern pin group '''''
    If UCase(glb_TesterType) = UCase("UltraFLEXplus") Then
        TheHdw.Digital.Patgen.ScanBurstEnabled = False
        TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = False
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Harvest_DigSrc(Harv_DigSrc As String, pat As Variant)
On Error GoTo errHandler
    Dim funcName As String: funcName = "Harvest_DigSrc"
    Dim sSrcSigName As String
    Dim tempVarArray As Variant
    Dim PatCount As Long
    Dim tempHarvDSSCHeader As String
    Dim tempHarvDSSCPin As New PinList
    Dim tempHarvDsp As New DSPWave
    Dim tempHarvDsp_SampleSize As Long
    Dim PatArray_absolute() As String
    
    tempHarvDSSCHeader = Split(Harv_DigSrc, ":")(0)
    tempHarvDSSCPin.value = Split(Harv_DigSrc, ":")(1)
    Call Harvest_CreateDigSrc(CStr(pat), tempHarvDSSCHeader, tempHarvDsp, tempHarvDsp_SampleSize)
    Call GetPatsFromPatSets(CStr(pat), PatArray_absolute(), PatCount, True)
    
    tempVarArray = TheHdw.DSSC.pins(tempHarvDSSCPin).Pattern(PatArray_absolute(0)).Source.Labels.list
    
    sSrcSigName = tempVarArray(0)
    If sSrcSigName = "" Then sSrcSigName = "FUNC_SRC_Harv"
    
    Call SetupDigSrcDspWave(PatArray_absolute(0), tempHarvDSSCPin, sSrcSigName, tempHarvDsp_SampleSize, tempHarvDsp)
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Harvest", funcName)
    If AbortTest Then Exit Function Else Resume Next
End Function

''20240124: Added for new user function usage for MFSTP
Public Function Parsing_UserFunction_Sheet() As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Parsing_UserFunction_Sheet"
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim sheetName As Variant, sheetnames() As String
    Dim ws As Worksheet
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim vTableContent As Variant
    Dim vTableContent_Transpose As Variant
    Dim HeaderRowArr() As String
    Dim HeaderCnt As Long
    Dim PatValue As String
    Dim BitValue As String
    Dim BitArr() As Long
    Dim BitArrCnt As Long
    Dim BitArr_StartRow() As Long
    Dim BitArr_EndRow() As Long
    Dim BitArr_SkipRow() As Long
    Dim foundSkipRow As Boolean
    Dim HarvDSSC_Value As String
    Dim HarvDSSC_StrArr() As String
    Dim HarvDSSC_FlagDic As New Dictionary
    Dim HarvDSSC_FlagArr() As String
    Dim HarvDSSC_FlagArrCnt As Long
    Dim HarvDicKey  As String
    
    Dim TempCnt As Long
    Dim TempStrArr() As String
    
    Const DefaultSheetName = "HARVMappingTable"
    Const Col1_Pat = 1
    Const Col2_Bit = 2
    Const Col4_DSSC = 4
    
    sheetnames = TheExec.job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)

    Dim foundSheetName As String
    For Each sheetName In sheetnames
        If sheetName Like DefaultSheetName Then
            foundSheetName = sheetName
            Exit For
        End If
    Next
    
    Static bIsUFMappingTableParsed As Boolean
    If bIsUFMappingTableParsed = False And foundSheetName <> "" Then
        gHarvDic_MappingTabble_SrcStr.RemoveAll
        gHarvDic_MappingTabble_FlagContained.RemoveAll
        gHarvDic_MappingTabble_BitNum.RemoveAll
        DigSrcPatternDict.RemoveAll
        
        'Each "UF_DigSrc_*" sheet
'        For Each sheetName In foundSheetNames
        Set ws = Sheets(foundSheetName)
        ws.Activate
        MaxRow = ws.UsedRange.Rows.Count
        maxcol = ws.UsedRange.Columns.Count
        vTableContent = ws.range(ws.Cells(1, 1), ws.Cells(MaxRow, maxcol)).value
        vTableContent_Transpose = Application.WorksheetFunction.Transpose(vTableContent)
        
        'record the header location
        HeaderCnt = 0
        For i = 1 To MaxRow
            If UCase(vTableContent_Transpose(Col1_Pat, i)) = UCase("PatternName") Then
                ReDim Preserve HeaderRowArr(HeaderCnt)
                ReDim Preserve BitArr_StartRow(HeaderCnt)
                ReDim Preserve BitArr_EndRow(HeaderCnt)
                
                HeaderRowArr(HeaderCnt) = i
                BitArr_StartRow(HeaderCnt) = i + 1
                If HeaderCnt <> 0 Then
                    BitArr_EndRow(HeaderCnt - 1) = i - 1
                End If
                HeaderCnt = HeaderCnt + 1
                
                foundSkipRow = False
            End If
            
            '20240124: New DigSrc sheet format. Find skip columns after any header was found
            If HeaderCnt > 0 And foundSkipRow = False Then
                If UCase(vTableContent_Transpose(Col2_Bit, i)) = "SKIP" Or vTableContent_Transpose(Col2_Bit, i) = "" Then
                    ReDim Preserve BitArr_SkipRow(HeaderCnt - 1)
                    BitArr_SkipRow(HeaderCnt - 1) = i
                    foundSkipRow = True
                End If
            End If
        Next i
        BitArr_EndRow(HeaderCnt - 1) = i - 1
        
        For i = 0 To UBound(HeaderRowArr)
            BitArrCnt = 0
            'PatArrCnt = 0
'                For j = BitArr_StartRow(i) To BitArr_EndRow(i)
            For j = BitArr_StartRow(i) To BitArr_SkipRow(i) - 1
                'only consider from 0 to XXX, from small to big, LSB to MSB
                BitValue = vTableContent_Transpose(Col2_Bit, j)
                If BitValue <> "" Then
                    If IsNumeric(BitValue) Then
                        ReDim Preserve BitArr(BitArrCnt)
                        BitArr(BitArrCnt) = j
                    Else
                        TempStrArr = Split(BitValue, "-")
                        TempCnt = Abs(TempStrArr(0) - TempStrArr(1))
                        ReDim Preserve BitArr(BitArrCnt + TempCnt)
                        For k = BitArrCnt To UBound(BitArr)
                            BitArr(k) = j
                        Next k
                        BitArrCnt = BitArrCnt + TempCnt
                    End If
                    BitArrCnt = BitArrCnt + 1
                End If
                
                '20240124: Save digsrc patterns into dictionary
                PatValue = UCase(vTableContent_Transpose(Col1_Pat, j))
                If PatValue <> "" And PatValue <> "END" Then
                    If DigSrcPatternDict.Exists(PatValue) Then
                        TheExec.Datalog.WriteComment "<Error> Harvest_DigSrc_MappingTableParsing : duplicate Pat Name."
                        Exit Function
                    End If
                    
                    ''-----Save Dig Src Patterns into dictionary----
                    DigSrcPatternDict.Add PatValue, vbNullString
                    ''----------------------------------------------
                    
                End If
            Next j
            
            For j = Col4_DSSC To maxcol
                Do  'Added for skipping for loop without GoTo
                    If UCase(vTableContent_Transpose(j, BitArr_SkipRow(i))) = "SKIP" Then
                        Exit Do     'Means continue
                    End If
                    
                    If UCase(vTableContent_Transpose(j, HeaderRowArr(i))) = "END" Then
                        Exit For
                    End If
                    
                    If vTableContent_Transpose(j, HeaderRowArr(i)) <> "" Then        ''X3G1
                        ReDim HarvDSSC_StrArr(UBound(BitArr))
                        ReDim HarvDSSC_FlagArr(0)
                        HarvDSSC_FlagArrCnt = 0
                        HarvDSSC_FlagDic.RemoveAll
                        For k = 0 To UBound(BitArr)
                            HarvDSSC_Value = vTableContent_Transpose(j, BitArr(k))
                            HarvDSSC_StrArr(k) = HarvDSSC_Value
                            If IsNumeric(HarvDSSC_Value) = False Then
                                HarvDSSC_Value = Replace(HarvDSSC_Value, "!", "")
                                If HarvDSSC_FlagDic.Exists(HarvDSSC_Value) = False Then
                                    HarvDSSC_FlagDic.Add HarvDSSC_Value, vbNullString
                                    ReDim Preserve HarvDSSC_FlagArr(HarvDSSC_FlagArrCnt)
                                    HarvDSSC_FlagArr(HarvDSSC_FlagArrCnt) = HarvDSSC_Value
                                    HarvDSSC_FlagArrCnt = HarvDSSC_FlagArrCnt + 1
                                End If
                            End If
                        Next k
                        
                        ''---------------Save Dig Src Bit info into dictionary--------------
                        HarvDicKey = vTableContent_Transpose(j, HeaderRowArr(i))
                        gHarvDic_MappingTabble_SrcStr.Add HarvDicKey, UCase(Join(HarvDSSC_StrArr, ","))
                        gHarvDic_MappingTabble_FlagContained.Add HarvDicKey, UCase(Join(HarvDSSC_FlagArr, ","))
                        gHarvDic_MappingTabble_BitNum.Add HarvDicKey, UBound(BitArr) + 1
                        ''------------------------------------------------------------------
                    End If
                Loop While False    'Added for skipping for loop without GoTo
            Next j
        Next i
        bIsUFMappingTableParsed = True
'        Next sheetName
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function HarvestingMainProcedure(patset() As String, ReportResult As PFType, TL_C_YES As Long, ResultMode As tlResultMode, _
            ConcurrentMode As tlPatConcurrentMode, ByRef SCAN_Site_Blooean, Optional ApplyVoltageFromBinCut As String = vbNullString, _
            Optional Harv_FailFlag As String = vbNullString, Optional HarvestPinGrpOtherFail As String = vbNullString)
    
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "HarvestingMainProcedure"
    Dim pat As Variant
    Dim inst_info As Instance_Info
    Dim instrumentUtility As New Instrument_Utility
    Dim i As Long
    Dim RowCnt As Long
    Dim ATPG_Pin_Table_Row As Long
    Dim Bool_CheckInitPat As Boolean
    Dim Search_ATPG_Harvest_Flag As Boolean
    
    Dim strAry_PathSplit() As String
    Dim strAry_PatNameSplit() As String
    Dim HarvConditionArr() As String
    Dim HarvFailFlagArr() As String
    Dim CustHarvCondition() As Boolean
    
    Dim sBool_PatternPass As New SiteBoolean
    
    If InStr(UCase(Harv_FailFlag), UCase("HarvestPinFlag_Table")) <> 0 Then
        Call CheckHarvestingOrMaskFromTable(Harv_FailFlag, EnableCoreHarvest, EnableCoreMask)
    ElseIf Harv_FailFlag <> "" Then
        Call Harvest_FailFlagSplit(Harv_FailFlag, HarvConditionArr, HarvFailFlagArr, CustHarvCondition)
        HarvFailCnt = UBound(HarvFailFlagArr) + 1
    End If
                    
    Search_ATPG_Harvest_Flag = False
    For Each pat In patset
                         
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'SycnUp setting
                         
        If ApplyVoltageFromBinCut <> "" Then
            For i = 0 To UBound(selsramLogicPingroup)
                If UCase(selsramLogicPingroup(i)) <> "PRESERVED" And UCase(selsramLogicPingroup(i)) <> "RESERVED" Then
                    If (TheHdw.DCVS.pins(selsramLogicPingroup(i)).Voltage.Output = tlDCVSVoltageAlt) Then
                        inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                        Exit For
                    End If
                End If
            Next i
                            
            'SycnUp
            If Flag_SyncUp_DCVS_Output_enable Then
                Call SyncUp_DCVS_Output(inst_info.p_mode, inst_info.currentDcvsOutput, SyncUp_PowerPin_Group) '''This is to sync up logic powers and sram powers on the same DCVS output (for TD testing)
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Start Search_ATPG_Harvest_Flag From HarvestPinFlag_Table
        If InStr(UCase(Harv_FailFlag), UCase("HarvestPinFlag_Table")) <> 0 Then
            strAry_PathSplit = Split(LCase(pat), "\")       '' Pattern name
            strAry_PatNameSplit = Split(LCase(strAry_PathSplit(UBound(strAry_PathSplit))), "_")     '' "in" or "pl"
            'Only payload pat need to check harv pin
            If strAry_PatNameSplit(3) Like "*pl*" Then
                'Search ATPG_HarvPinFlagMapping to find the correct Harvest Pin Input
                For RowCnt = 0 To UBound(HarvPinFlagMapping)
                    If LCase(strAry_PathSplit(UBound(strAry_PathSplit))) Like LCase(HarvPinFlagMapping(RowCnt).Pattern) Then    '' ex: "*CCC0*" matched
                        HarvConditionArr = HarvPinFlagMapping(RowCnt).HarvPinGrpConditionArr
                        HarvFailFlagArr = HarvPinFlagMapping(RowCnt).HarvPinGrpCondFailFlagArr
    '                   HarvestPinGrpOtherFail = HarvPinFlagMapping(RowCnt).OtherPinFailFlag
    '                   OtherPinBinOutStage = HarvPinFlagMapping(RowCnt).OtherPinBinOutStage
                        HarvFailCnt = UBound(HarvFailFlagArr) + 1
                                    
                        Search_ATPG_Harvest_Flag = True
                        ATPG_Pin_Table_Row = RowCnt 'Record Row number
                        Exit For
                    
                    End If
                Next RowCnt
            End If
        End If
        'End Search_ATPG_Harvest_Flag From HarvestPinFlag_Table
        
        If TheExec.TesterMode = testModeOffline Then
            Call ATPG_offline(CStr(pat), ResultMode)
        Else
            If gl_bTTRDisableAlarm = False Then     'T-Col TTR approve by Si -- 230413
                TheHdw.Alarms.Check
            End If
                            
            'If m_ScanLogging = True Then
            '    Call thehdw.Patterns(CStr(pat)).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)
            'Else
                Glb_Bool_CheckSFCInitPat = False 'Initial Flag for each pattern
                Call CheckInitPat(CStr(pat), Bool_CheckInitPat)
                Glb_Bool_CheckSFCInitPat = Bool_CheckInitPat 'Set Flag status, TRUE means Init Pattern, False means PL Pattern
                If Bool_CheckInitPat = False Then
                    If InStr(UCase(Harv_FailFlag), UCase("HarvestPinFlag_Table")) <> 0 Then
                        If Search_ATPG_Harvest_Flag = True Then
                            Call Harvest_CMEM_InitSetup
                        Else
                            TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                            TheExec.Datalog.WriteComment "Harvesting PinGorup did not exist on HarvestPinFlag_Table, Could not run Harvest_CMEM_InitSetup, Please check it."
    '                       TheExec.Datalog.WriteComment "Search_ATPG_Harvest_Flag = False, Turn TRUE strGlb_Flag_HarvestPinLose_Fail_Stop Flag."
    '                       For Each site In TheExec.sites
    '                           TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestPinLose_Fail_Stop) = logicTrue
    '                       Next site
                        End If
                    ElseIf Harv_FailFlag <> "" Or (glb_SFC_Scan_Check = True) Then
                        Call Harvest_CMEM_InitSetup
                    End If
                End If
                
                If Bool_CheckInitPat = False Then
                    If InStr(UCase(Harv_FailFlag), UCase("HarvestPinFlag_Table")) <> 0 Then
                        If Search_ATPG_Harvest_Flag = True Then
                                            
                            Call Enable_DisableComparePinFromFuse(HarvConditionArr, HarvFailFlagArr)
                                            
                            If EnableCoreMask = True Then
                                'ATPG Harvest Pin Flag Table
                                Call Enable_PinMaskFeature(HarvConditionArr, HarvFailFlagArr)
                            End If
                        Else
                            TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                            TheExec.Datalog.WriteComment "Harvesting PinGorup did not exist on HarvestPinFlag_Table, Could not run Enable_PinMaskFeature, Please check it."
 '                          TheExec.Datalog.WriteComment "Search_ATPG_Harvest_Flag = False, Turn TRUE strGlb_Flag_HarvestPinLose_Fail_Stop Flag."
 '                          For Each site In TheExec.sites
 '                              TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestPinLose_Fail_Stop) = logicTrue
 '                          Next site
                        End If
                    ElseIf Harv_FailFlag <> "" Then
                        Call HarvCustomFeature(HarvConditionArr, HarvFailFlagArr, CustHarvCondition)
                    End If
                End If
                
                Call TheHdw.patterns(CStr(pat)).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)
        End If
        
        If Flag_HarvPinFlag_Mapping_Table_Parsed = True Then
            If Bool_CheckInitPat = False Then
                HarvPinsFailCnt = 0  'Initial HarvPinsFailCnt  to 0
                Call checkIfHarvestingPinFail(CStr(pat), Search_ATPG_Harvest_Flag, EnableCoreHarvest, HarvFailFlagArr, HarvConditionArr, HarvestPinGrpOtherFail, CustHarvCondition, Harv_FailFlag)
                Call checkIfOtherPinFail(HarvPinsFailCnt, HarvestPinGrpOtherFail)
            Else
                sBool_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
                For Each site In TheExec.sites
                    If sBool_PatternPass(site) = False Then
                        TheExec.sites.item(site).FlagState(HarvestPinGrpOtherFail) = logicTrue
                    End If
                Next site
            End If
        Else
            'PrintFailPat_Flag = TheExec.Datalog.Setup.DatalogSetup.SelectSetupFile
            If TheExec.enableWord("PatternFailInfo") = True And TheHdw.Digital.hram.size <> 0 Then
                Call Printing_StandalonePat(CStr(pat), patset)
            End If
        End If
        
        If Bool_CheckInitPat = False Then
            If InStr(UCase(Harv_FailFlag), UCase("HarvestPinFlag_Table")) <> 0 Then
                If Search_ATPG_Harvest_Flag = True Then
                    Call Disable_PinMaskFeature(HarvConditionArr, HarvFailFlagArr)
                Else
                    TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                    TheExec.Datalog.WriteComment "Harvesting PinGorup did not exist on HarvestPinFlag_Table, Could not run Disable_PinMaskFeature, Please check it."
'                   TheExec.Datalog.WriteComment "Search_ATPG_Harvest_Flag = False, Turn TRUE strGlb_Flag_HarvestPinLose_Fail_Stop Flag."
'                   For Each site In TheExec.sites
'                       TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestPinLose_Fail_Stop) = logicTrue
'                   Next site
                End If
            ElseIf Harv_FailFlag <> "" Then
                Call Harvest_CMEM_Stop(HarvConditionArr, CustHarvCondition, HarvFailFlagArr)
            ElseIf glb_SFC_Scan_Check = True Then
                Call SFC_CMEM_Stop
            Else
            End If
        End If

        Search_ATPG_Harvest_Flag = False
    Next pat
    
    Exit Function
                                                                                                                                                                                                                                                               
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CheckHarvestingOrMaskFromTable(Harv_FailFlag As String, ByRef EnableCoreHarvest As Boolean, ByRef EnableCoreMask As Boolean)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "CheckHarvestingOrMaskFromTable"
    Dim h As Long
    Dim Harvest_Argument_Info() As String
    Dim Harvest_enableword_Info() As String
    
    
    
        If UCase(Harv_FailFlag) = UCase("HarvestPinFlag_Table") Then
        'just keep the original Harvest_FailFlagSplit, HarvConditionArr & HarvFailFlagArr will get from HarvestPinFlag_Table when pattern loop match the related pattern key word, 20230626 Jim
        'Read Enableword "EnableCoreHarvest" & "EnableCoreMask" status when Get Harvest pin/fail flag information from HarvestPinFlag_Table.
        'Check whether to Enable Harvest Flag, if EnableWord "EnableCoreHarvest" = TRUE, and Harvest pin group test fail, then IGXL will Set Harvest Flag to TRUE
            If TheExec.flow.enableWord("EnableCoreHarvest") = True Then
                EnableCoreHarvest = True
            Else
                EnableCoreHarvest = False
            End If
                            
            'Check whether to Enable DisableCompare, if EnableWord "EnableCoreMask" = TRUE, and the new Harvest Fail at current stage, then IGXL will DisableCompare the related Harvest pin
            If TheExec.flow.enableWord("EnableCoreMask") = True Then
                EnableCoreMask = True
            Else
                EnableCoreMask = False
            End If
                        
            ElseIf InStr(UCase(Harv_FailFlag), UCase("EnableCoreHarvest")) <> 0 And InStr(UCase(Harv_FailFlag), UCase("EnableCoreMask")) <> 0 And InStr(UCase(Harv_FailFlag), ":") <> 0 And InStr(UCase(Harv_FailFlag), ";") <> 0 Then
                            
                Harvest_Argument_Info = Split(UCase(Harv_FailFlag), ";")
                            
                If UBound(Harvest_Argument_Info) = 2 Then
                                
                    For h = 1 To UBound(Harvest_Argument_Info)
                                    
                        Harvest_enableword_Info = Split(Harvest_Argument_Info(h), ":")
                                    
                        If UBound(Harvest_enableword_Info) = 1 Then
                            'Check whether to Enable Harvest Flag, if "EnableCoreHarvest" = TRUE, and Harvest pin group test fail, then IGXL will Set Harvest Flag to TRUE
                            If UCase(Harvest_enableword_Info(0)) = UCase("EnableCoreHarvest") Then
                                If UCase(Harvest_enableword_Info(1)) = "TRUE" Then
                                    EnableCoreHarvest = True
                                ElseIf UCase(Harvest_enableword_Info(1)) = "FALSE" Then
                                    EnableCoreHarvest = False
                                Else
                                    'Print error message to Datalog,  "Harv_FailFlag" Format is incorrect
                                    TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                                    TheExec.Datalog.WriteComment "Harv_FailFlag Format is incorrect, [EnableCoreHarvest] Boolean Setting Wrong, please check it."
                                    TheExec.Datalog.WriteComment "Harv_FailFlag : " & Harv_FailFlag
                                End If
                            'Check whether to Enable DisableCompare, if "EnableCoreMask" = TRUE, and the new Harvest Fail at current stage, then IGXL will DisableCompare the related Harvest pin
                            ElseIf UCase(Harvest_enableword_Info(0)) = UCase("EnableCoreMask") Then
                                If UCase(Harvest_enableword_Info(1)) = "TRUE" Then
                                    EnableCoreMask = True
                                ElseIf UCase(Harvest_enableword_Info(1)) = "FALSE" Then
                                    EnableCoreMask = False
                                Else
                            'Print error message to Datalog,  "Harv_FailFlag" Format is incorrect
                                    TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                                    TheExec.Datalog.WriteComment "Harv_FailFlag Format is incorrect, [EnableCoreMask] Boolean Setting Wrong, please check it."
                                    TheExec.Datalog.WriteComment "Harv_FailFlag : " & Harv_FailFlag
                                End If
                            Else
                            'Print error message to Datalog,  "Harv_FailFlag" Format is incorrect
                                    TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                                    TheExec.Datalog.WriteComment "Harv_FailFlag Format is incorrect, [EnableCoreHarvest] or [EnableCoreMask] Setting Wrong, please check it."
                                    TheExec.Datalog.WriteComment "Harv_FailFlag : " & Harv_FailFlag
                            End If
                        Else
                            'Print error message to Datalog,  "Harv_FailFlag" Format is incorrect
                            TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                            TheExec.Datalog.WriteComment "Harv_FailFlag Format is incorrect, [:] Setting Wrong, please check it."
                            TheExec.Datalog.WriteComment "Harv_FailFlag : " & Harv_FailFlag
                        End If
                    Next h
                Else
                    'Print error message to Datalog,  "Harv_FailFlag" Format is incorrect
                    TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                    TheExec.Datalog.WriteComment "Harv_FailFlag Format is incorrect, [;] Setting Wrong, please check it."
                    TheExec.Datalog.WriteComment "Harv_FailFlag : " & Harv_FailFlag
                End If
            Else
                'Print error message to Datalog,  "Harv_FailFlag" Format is incorrect
                TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
                TheExec.Datalog.WriteComment "Harv_FailFlag Format is incorrect, please check it."
                TheExec.Datalog.WriteComment "Harv_FailFlag : " & Harv_FailFlag
            End If

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function checkIfHarvestingPinFail(pat As String, Search_ATPG_Harvest_Flag As Boolean, EnableCoreHarvest As Boolean, _
            HarvFailFlagArr() As String, HarvConditionArr() As String, HarvestPinGrpOtherFail As String, CustHarvCondition() As Boolean, Optional Harv_FailFlag As String)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "checkIfHarvestingPinFail"
    
    Dim HarvCnt As Long
    Dim site As Variant
    'Dim Harv_FailFlag As String
    Dim Sbln_PatternPass As New SiteBoolean
    If InStr(UCase(Harv_FailFlag), UCase("HarvestPinFlag_Table")) <> 0 Then
        If Search_ATPG_Harvest_Flag = True Then
            If EnableCoreHarvest = True Then
                TheExec.Datalog.WriteComment "[ Turn on Harvest Fail Flag ] - Harvest_Pin_From_Table."
                'ATPG Harvest Pin Flag Table
                For HarvCnt = 0 To UBound(HarvFailFlagArr)
                    Call Harvest_Decision(HarvFailFlagArr(HarvCnt), HarvConditionArr(HarvCnt), CStr(pat))
                Next HarvCnt
            Else
                Sbln_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
                If Sbln_PatternPass.Any(False) Then
                    For Each site In TheExec.sites.Selected
                        If Sbln_PatternPass(site) = False Then
                            TheExec.Datalog.WriteComment "Site" & site & " EnableCoreHarvest = False & Pattern fail, Turn TRUE HarvestPinGrpOtherFail Flag."
                            TheExec.sites.item(site).FlagState(HarvestPinGrpOtherFail) = logicTrue 'Because of Search_ATPG_Harvest_Flag = True, so Turn True HarvestPinGrpOtherFail Flag to Binout
                        End If
                    Next site
                End If
            End If
                            
        Else
            TheExec.Datalog.WriteComment "Instance Name: " & glb_TestInstance
            TheExec.Datalog.WriteComment "Harvesting PinGorup did not exist on HarvestPinFlag_Table, Could not run Harvest_Decision, Please check it."
'                              TheExec.Datalog.WriteComment "Search_ATPG_Harvest_Flag = False, Turn TRUE strGlb_Flag_HarvestPinLose_Fail_Stop Flag."
'                              For Each site In TheExec.sites
'                                  TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestPinLose_Fail_Stop) = logicTrue
'                              Next site
        End If
    Else
        For HarvCnt = 0 To UBound(HarvFailFlagArr)
            Call Harvest_Decision(HarvFailFlagArr(HarvCnt), HarvConditionArr(HarvCnt), CStr(pat), CustHarvCondition(HarvCnt))
        Next HarvCnt
    End If

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function checkIfOtherPinFail(sl_HarvPinsFailCnt As SiteLong, s_HarvestPinGrpOtherFail As String)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "checkIfOtherPinFail"
    Dim site As Variant
    Dim TempDigitalPin() As String
    For Each site In TheExec.sites
'''''''  If HarvFailCnt(Site) = 0 Then ''''' nop for binout NonHarv Pin fail 20230216
'''''''     TheExec.sites.item(Site).FlagState(s_HarvestPinGrpOtherFail) = logicTrue
'''''''  ElseIf HarvOtherFailCnt(Site) <> 0 Then
'''''''     TempDigitalPin = TheHdw.Digital.FailedPins(Site)
'''''''     If UBound(TempDigitalPin) + 1 > HarvOtherFailCnt(Site) Then
'''''''        heExec.sites.item(Site).FlagState(s_HarvestPinGrpOtherFail) = logicTrue
'''''''     End If
'''''''  End If
        TempDigitalPin = TheHdw.Digital.FailedPins(site)
        If HarvPinsFailCnt(site) <> 0 Then
            If UBound(TempDigitalPin) + 1 > HarvPinsFailCnt(site) Then
                TheExec.sites.item(site).FlagState(s_HarvestPinGrpOtherFail) = logicTrue
            ElseIf UBound(TempDigitalPin) + 1 = HarvPinsFailCnt(site) Then
                TheExec.Datalog.WriteComment "HarvPinsFailCnt = All pattern fail pin count do nothing!"
            Else
                TheExec.Datalog.WriteComment "HarvPinsFailCnt Large than All pattern fail pin count"
                TheExec.sites.item(site).FlagState(s_HarvestPinGrpOtherFail) = logicTrue
            End If
        ElseIf UBound(TempDigitalPin) + 1 <> 0 And HarvPinsFailCnt(site) = 0 Then
            TheExec.sites.item(site).FlagState(s_HarvestPinGrpOtherFail) = logicTrue
            ElseIf UBound(TempDigitalPin) + 1 = 0 And HarvPinsFailCnt(site) = 0 Then
                TheExec.Datalog.WriteComment "Pattern Pass without any fail pin" 'Add comment for All Pass
            Else
                TheExec.Datalog.WriteComment "Please Confrim Harvesting/NonHarvesting Pins Fail Status"
                TheExec.sites.item(site).FlagState(s_HarvestPinGrpOtherFail) = logicTrue ' Binout Other case
            End If
    Next site
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
