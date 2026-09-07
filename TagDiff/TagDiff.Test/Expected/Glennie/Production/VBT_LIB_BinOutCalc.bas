Attribute VB_Name = "VBT_LIB_BinOutCalc"


Private Type FlagState
    strFlagName As String
    strPassOrFail As String
    strFailAndNoTestFlag As String
    strFlagCheckInfor As String
End Type
Private Type FlagInfor
    astrFlag() As FlagState
    'astrHVflag() As FlagState
    'astrLVflag() As FlagState
    iSite As Integer
    lFailCounts As Long

End Type
Private Type SiteInfor
    astrSiteVal() As FlagInfor
    strBlockFlagName As String
End Type

Private Type BlockFlag
    astrBlockFlag() As SiteInfor
End Type

Private Type BlockFailNoTestFlag
    astrBlockFlag() As FlagInfor
End Type

Private Type BinOutInfor
    astrBlackName() As String
    strBinFlag As String
End Type
Private Type BinFlag
    astrBinFlagName() As BinOutInfor
End Type

Public tyBinOutInfor As BinFlag
Public blReadBinOutFlag As Boolean
Public tyAllBlockFlagInfor As BlockFlag
Public tyAllFailAndNoTestBlockFlagInfor As BlockFailNoTestFlag
'update command to avoid IGXL version switching error
'If theexec.SoftwareVersion Like "*9.10*" Then
    Public Declare PtrSafe Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
'Else
  '  Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
'End If
Public strAllFlagArr() As String       ''20160504 ---webster
Public gl_iExistSiteCount As Integer
'Private strFlagInforPrint() As String


Public Function BinOutCalc_CleanFlagStatus(ScanMbistSheet As String, HardIPSheet As String, BinOutSheet As String, Optional EnableClearFlag As Boolean = True) As Long

' V: not test logicClear => fail:logicTrue       ; have to be tested       , if flag is clear and setting "V", same with test fail
' C: not test logicClear => pass:logicFalse      ; Do not have to be tested, if flag is clear and setting "C", same with test pass
' "": do not care: not Binning
' S: the same with "V", only use in FT
' tyBinOutInfor.astrBinFlagName(BinFlagNum - 1)       : BinOutTable_CP1/BinOutTable_CP2/BinOutTable_FT1
' tyAllBlockFlagInfor.astrBlockFlag (iBlockFlagIdx)   : BinOutCalcScanMbistTable
' Flag        : And (Flag-CPU_Chain_SA_all)
' Fail        : at least one flag fail
' FlagOr      : Or (FlagOr-CPU_Chain_SA_all)
' Pass        : at least one flag pass
' EnableHVorLV: true=>HV or LV pass, The test condition is pass

    On Error GoTo errHandler
    Dim funcName As String:: funcName = "BinOutCalc_CleanFlagStatus"

    Dim BinOutCalcSheet As Worksheet

    Dim BlockColumnMax As Long
    Dim blockNum As Long
    Dim BinFlagNum As Long
    Dim BlockContext As String
    Dim TestInstRowMax As Long
    Dim TestInstRowStart As Long
    Dim iColumnIdx As Integer
    Dim index As Long
    Dim flagName As String
    Dim site As Variant
    Dim tyBinOutInfor1 As BinFlag
    Dim iBlockIdx As Integer
    Dim lFlagCount As Long   ''add 20160504 ---webster
    Dim tyEveBlockFlagInfor As SiteInfor
     Dim tyEveBlockFlagInfor1 As SiteInfor   ''20160510
'    Dim tyEveFailAndNoTestBlockFlagInfor As FlagInfor
    Dim iSiteIdx As Integer
    Dim iFlagIdx() As Integer
    Dim strFlagCheck As String
    Dim iBlockFlagIdx As Integer
    Dim strBinBlockFlagCell As String

    If blReadBinOutFlag = False Then
        
        lFlagCount = 0   ''add 20160504 ---webster
        gl_iExistSiteCount = TheExec.sites.Existing.Count  '' Can put onProgramValidate
        ReDim tyEveBlockFlagInfor.astrSiteVal(gl_iExistSiteCount - 1)
        ReDim iFlagIdx(gl_iExistSiteCount - 1)
    
        Dim tyAllBlockFlagInfor1 As BlockFlag
        Dim tyAllFailAndNoTestBlockFlagInfor1 As BlockFailNoTestFlag
        tyAllBlockFlagInfor = tyAllBlockFlagInfor1
        tyAllFailAndNoTestBlockFlagInfor = tyAllFailAndNoTestBlockFlagInfor1
        ReDim tyAllBlockFlagInfor.astrBlockFlag(100)
        ReDim tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(100)

        tyBinOutInfor = tyBinOutInfor1
        
        TestInstRowStart = 2
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' read ScanMbist BinOutCalc sheet
        Set BinOutCalcSheet = Worksheets(ScanMbistSheet)
        BlockColumnMax = BinOutCalcSheet.Cells(1, Columns.Count).End(xlToLeft).Column    '//x max
        For iColumnIdx = 1 To BlockColumnMax                                             '//x loop
            BlockContext = Trim(BinOutCalcSheet.Cells(1, iColumnIdx).value)              '//x index
            If UCase(BlockContext) Like "*FLAG*" Then
                tyEveBlockFlagInfor.strBlockFlagName = BlockContext
                TestInstRowMax = BinOutCalcSheet.Cells(Rows.Count, iColumnIdx).End(xlUp).Row    '//y max
                For index = TestInstRowStart To TestInstRowMax                                  '//y loop
                    flagName = Trim(BinOutCalcSheet.Cells(index, iColumnIdx))                   '//y value
                    strFlagCheck = Trim(BinOutCalcSheet.Cells(index, iColumnIdx + 1))           '//V or C or ""
                     ''save all flag    add 20160504 ---webster
                    If flagName <> "" Then
                         ReDim Preserve strAllFlagArr(lFlagCount)
''                        If strFlagCheck <> "" Then '' need update , only reinit care flag
                            strAllFlagArr(lFlagCount) = flagName
                            lFlagCount = lFlagCount + 1
                        If strFlagCheck <> "" Then                                              '//check V or C or S
                            For iSiteIdx = 0 To gl_iExistSiteCount - 1                          '//site loop
                                 ReDim Preserve tyEveBlockFlagInfor.astrSiteVal(iSiteIdx).astrFlag(iFlagIdx(iSiteIdx))
                                 tyEveBlockFlagInfor.astrSiteVal(iSiteIdx).astrFlag(iFlagIdx(iSiteIdx)).strFlagName = flagName                '//flag name
                                 tyEveBlockFlagInfor.astrSiteVal(iSiteIdx).astrFlag(iFlagIdx(iSiteIdx)).strFlagCheckInfor = strFlagCheck      '//flag name state(V or C or S)
                                 iFlagIdx(iSiteIdx) = iFlagIdx(iSiteIdx) + 1
                            Next iSiteIdx

                        End If
                    End If          ''If FlagName
                Next index

                If iBlockFlagIdx >= 100 Then
                    ReDim Preserve tyAllBlockFlagInfor.astrBlockFlag(iBlockFlagIdx)
                    ReDim Preserve tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(iBlockFlagIdx)
                End If
                tyAllBlockFlagInfor.astrBlockFlag(iBlockFlagIdx) = tyEveBlockFlagInfor   '//tyAllBlockFlagInfor.astrBlockFlag(iBlockFlagIdx) store mbist sheet , iBlockFlagIdx: x number
                tyEveBlockFlagInfor = tyEveBlockFlagInfor1                               '//clean tyEveBlockFlagInfor
                iBlockFlagIdx = iBlockFlagIdx + 1
                ReDim tyEveBlockFlagInfor.astrSiteVal(gl_iExistSiteCount - 1)
                ReDim iFlagIdx(gl_iExistSiteCount - 1) '' reset
            End If '' If UCase(BlockContext) Like "*FLAG*"
        Next iColumnIdx
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' read  HardIPSheet  BinOutCalc sheet
        Set BinOutCalcSheet = Worksheets(HardIPSheet)
        BlockColumnMax = BinOutCalcSheet.Cells(1, Columns.Count).End(xlToLeft).Column
        For iColumnIdx = 1 To BlockColumnMax
            BlockContext = Trim(BinOutCalcSheet.Cells(1, iColumnIdx).value)
            If UCase(BlockContext) Like "*FLAG*" Then
                tyEveBlockFlagInfor.strBlockFlagName = BlockContext
                TestInstRowMax = BinOutCalcSheet.Cells(Rows.Count, iColumnIdx).End(xlUp).Row
                For index = TestInstRowStart To TestInstRowMax
                    flagName = Trim(BinOutCalcSheet.Cells(index, iColumnIdx))
                    strFlagCheck = UCase(Trim(BinOutCalcSheet.Cells(index, iColumnIdx + 1)))
                     ''save all flag    add 20160504 ---webster
                    If flagName <> "" Then
                        ReDim Preserve strAllFlagArr(lFlagCount)
''                        If strFlagCheck <> "" Then '' need update , only reinit care flag
                            strAllFlagArr(lFlagCount) = flagName
                            lFlagCount = lFlagCount + 1
                        If strFlagCheck <> "" Then
                            For iSiteIdx = 0 To gl_iExistSiteCount - 1
                                 ReDim Preserve tyEveBlockFlagInfor.astrSiteVal(iSiteIdx).astrFlag(iFlagIdx(iSiteIdx))
                                 tyEveBlockFlagInfor.astrSiteVal(iSiteIdx).astrFlag(iFlagIdx(iSiteIdx)).strFlagName = flagName
                                 tyEveBlockFlagInfor.astrSiteVal(iSiteIdx).astrFlag(iFlagIdx(iSiteIdx)).strFlagCheckInfor = strFlagCheck
                                 iFlagIdx(iSiteIdx) = iFlagIdx(iSiteIdx) + 1
                            Next iSiteIdx

                        End If
                    End If          ''If FlagName
                Next index

                If iBlockFlagIdx >= 100 Then
                    ReDim Preserve tyAllBlockFlagInfor.astrBlockFlag(iBlockFlagIdx)
                    ReDim Preserve tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(iBlockFlagIdx)
                End If
                tyAllBlockFlagInfor.astrBlockFlag(iBlockFlagIdx) = tyEveBlockFlagInfor
                tyEveBlockFlagInfor = tyEveBlockFlagInfor1
                ReDim tyEveBlockFlagInfor.astrSiteVal(gl_iExistSiteCount - 1)
                iBlockFlagIdx = iBlockFlagIdx + 1
                ReDim iFlagIdx(gl_iExistSiteCount - 1) '' reset
            End If '' If UCase(BlockContext) Like "*FLAG*"
        Next iColumnIdx
            
        ReDim Preserve tyAllBlockFlagInfor.astrBlockFlag(iBlockFlagIdx - 1)
        ReDim Preserve tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(iBlockFlagIdx - 1)
        iBlockFlagIdx = 0
        lFlagCount = 0
        
        '''  ''read  BinOut sheet
        ''
        BinOutSheet = "BinOutTable"
        If currentJobName Like "*cp1*" Then
            BinOutSheet = BinOutSheet & "_CP1"
        ElseIf currentJobName Like "*cp2*" Then
            BinOutSheet = BinOutSheet & "_CP2"
        ElseIf currentJobName Like "*ft1*" Then
            BinOutSheet = BinOutSheet & "_FT1"
        ElseIf currentJobName Like "*ft2*" Then
            BinOutSheet = BinOutSheet & "_FT2"
        ElseIf currentJobName Like "*ft3*" Then
            BinOutSheet = BinOutSheet & "_FT3"
        Else
            BinOutSheet = BinOutSheet & "_CP1"
        End If
        
        Dim Sheet As Worksheet
        Dim IsSheetExist As Boolean
        IsSheetExist = False
        For Each Sheet In Worksheets
            If Sheet.name = BinOutSheet Then IsSheetExist = True: Exit For
        Next Sheet
        If IsSheetExist = False Then
            TheExec.Datalog.WriteComment ("<Warning> " & "Sheet: " & "(" & BinOutSheet & ")  do not exist, please check the sheet name or create the sheet!")
            Exit Function
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' BinOutTable_CP1/BinOutTable_CP2/BinOutTable_FT1
        Set BinOutCalcSheet = Worksheets(BinOutSheet)
        BlockColumnMax = BinOutCalcSheet.Cells(1, Columns.Count).End(xlToLeft).Column          '//x max
        ReDim Preserve tyBinOutInfor.astrBinFlagName(BlockColumnMax)
        For iColumnIdx = 1 To BlockColumnMax                                                   '//x loop
            BlockContext = Trim(BinOutCalcSheet.Cells(1, iColumnIdx).value)                    '//x index
            If UCase(BlockContext) Like "F_*" Then       'Save bin flag information        F_BIN -> F_
                TestInstRowMax = BinOutCalcSheet.Cells(Rows.Count, iColumnIdx).End(xlUp).Row   '//y max
                tyBinOutInfor.astrBinFlagName(BinFlagNum).strBinFlag = BlockContext
                If TestInstRowMax > 1 Then
                    ReDim Preserve tyBinOutInfor.astrBinFlagName(BinFlagNum).astrBlackName(TestInstRowMax - 2)
                    For index = TestInstRowStart To TestInstRowMax                              '//y loop
                        strBinBlockFlagCell = Trim(BinOutCalcSheet.Cells(index, iColumnIdx))    '//y value
                        If UCase(strBinBlockFlagCell) Like "*FLAG*" Then   ' update 20160427 webster
                            tyBinOutInfor.astrBinFlagName(BinFlagNum).astrBlackName(iBlockIdx) = strBinBlockFlagCell
                            iBlockIdx = iBlockIdx + 1
                        End If
                    Next index
                    If iBlockIdx = 0 Then iBlockIdx = 1
                    ReDim Preserve tyBinOutInfor.astrBinFlagName(BinFlagNum).astrBlackName(iBlockIdx - 1)
                    iBlockIdx = 0
                End If
                BinFlagNum = BinFlagNum + 1
            End If
        Next iColumnIdx

     
        ReDim Preserve tyBinOutInfor.astrBinFlagName(BinFlagNum - 1)       '//store BinOutTable_CP1 information
        BinFlagNum = 0
        blReadBinOutFlag = True

        '' remove duplicate flag   add 20160504 ---webster
        Call RemoveDeplicateFlag

    End If

    '' init flag status       add 20160504 ---webster
    '' update   need make sure not device condition or if
    If EnableClearFlag <> False Then
        Dim lFlagIdx As Long
        For Each site In TheExec.sites
            For lFlagIdx = 0 To UBound(strAllFlagArr)      '//BinOutCalcScanMbistTable : clear all flag state
                TheExec.sites.item(site).FlagState(strAllFlagArr(lFlagIdx)) = logicClear
            Next lFlagIdx
        Next site
    End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

' remove duplicate flag   add 20160504 ---webster
Private Sub RemoveDeplicateFlag()

If SafeArrayGetDim(strAllFlagArr) = 0 Then Stop
Dim strAllFlagArrTemp() As String
Dim lFlagCountTemp As Long
Dim lFlagIdx As Long
Dim lFlagTempIdx As Long
Dim iSameFlagCount As Integer

lFlagCountTemp = 0
iSameFlagCount = 0

For lFlagIdx = 0 To UBound(strAllFlagArr)
    For lFlagTempIdx = 0 To lFlagIdx
        If strAllFlagArr(lFlagIdx) = strAllFlagArr(lFlagTempIdx) Then
            iSameFlagCount = iSameFlagCount + 1
            If iSameFlagCount > 1 Then Exit For
        End If
    Next lFlagTempIdx
    If iSameFlagCount = 1 Then
        ReDim Preserve strAllFlagArrTemp(lFlagCountTemp)
        strAllFlagArrTemp(lFlagCountTemp) = strAllFlagArr(lFlagIdx)
        lFlagCountTemp = lFlagCountTemp + 1
    End If
    iSameFlagCount = 0
Next lFlagIdx

ReDim strAllFlagArr(lFlagCountTemp - 1)
strAllFlagArr = strAllFlagArrTemp

End Sub


''' Note: lFailCounter may not the true number of flag fail
''' Update for TTR  20160504  ---- webster
Public Function BinOutCalc_ReadBlockFlagState(Optional EnableHVorLV As Boolean = False) As Long


' V: not test logicClear => fail:logicTrue       ; have to be tested       , if flag is clear and setting "V", same with test fail
' C: not test logicClear => pass:logicFalse      ; Do not have to be tested, if flag is clear and setting "C", same with test pass
' "": do not care: not Binning
' S: the same with "V", only use in FT
' tyBinOutInfor.astrBinFlagName(BinFlagNum - 1)       : BinOutTable_CP1/BinOutTable_CP2/BinOutTable_FT1
' tyAllBlockFlagInfor.astrBlockFlag (iBlockFlagIdx)   : BinOutCalcScanMbistTable
' Flag        : And (Flag-CPU_Chain_SA_all)
' Fail        : at least one flag fail
' FlagOr      : Or (FlagOr-CPU_Chain_SA_all)
' Pass        : at least one flag pass
' EnableHVorLV: true=>HV or LV pass, The test condition is pass

    On Error GoTo errHandler
    Dim funcName As String:: funcName = "BinOutCalc_ReadBlockFlagState"

    Dim lFailCounter As Long
    Dim lFailAndNoTestCounter As Long
    Dim FlagStatus As tlLogicState
    Dim PreFlagStatus As tlLogicState
    Dim strTestFlagName As String
    Dim strTestFlagCheckInfor As String
    Dim strPreTestFlagCheckInfor As String
    Dim blIsFlagOrPass As Boolean
    Dim strTestFlagNameTemp As String
    Dim site As Variant
    Dim strBlockFlagName As String

    blIsFlagOrPass = False  '' make sure the default value is false  20160504 ----webster

''    ''' add for debug
   '' Call ReInitFlagStatus

    Dim lBlockFlagIdx As Long
    Dim lFlagIdx As Long
    
    For lBlockFlagIdx = 0 To UBound(tyAllBlockFlagInfor.astrBlockFlag)
         strBlockFlagName = tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).strBlockFlagName                           '//for BinOutCalcScanMbistTable


         ReDim Preserve tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrFlag((UBound(tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(0).astrFlag) + 1) * gl_iExistSiteCount - 1)
         For Each site In TheExec.sites
            For lFlagIdx = 0 To UBound(tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(site).astrFlag)     '//for BinOutCalcScanMbistTable
                strTestFlagName = tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(site).astrFlag(lFlagIdx).strFlagName         '' have restore bug
                strTestFlagCheckInfor = tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(site).astrFlag(lFlagIdx).strFlagCheckInfor

                FlagStatus = TheExec.sites.item(site).FlagState(strTestFlagName)
''''                ReDim Preserve tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrFlag(lFailAndNoTestCounter)
                '======================================================================================================================================================================
                With tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(lBlockFlagIdx)
                    If FlagStatus = logicTrue Then
'                        tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(Site).astrFlag(lFlagIdx).strPassOrFail = "F"
                        .astrFlag(lFailAndNoTestCounter).strFailAndNoTestFlag = site & " " & strBlockFlagName & " " & strTestFlagName & " Fail" & " " & strTestFlagCheckInfor
                        lFailAndNoTestCounter = lFailAndNoTestCounter + 1
                        lFailCounter = lFailCounter + 1                         'save to total fail counts for test item

                    ElseIf FlagStatus = logicFalse Then
'                        tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(Site).astrFlag(lFlagIdx).strPassOrFail = "P"
                        .astrFlag(lFailAndNoTestCounter).strFailAndNoTestFlag = site & " " & strBlockFlagName & " " & strTestFlagName & " Pass" & " " & strTestFlagCheckInfor
                        lFailAndNoTestCounter = lFailAndNoTestCounter + 1

                        If UCase(BlockContext) Like "*FLAGOR-*" Then   '//logic or
                           blIsFlagOrPass = True
                        End If

                    ElseIf FlagStatus = logicClear Then
'                        tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(Site).astrFlag(lFlagIdx).strPassOrFail = "X"
                        .astrFlag(lFailAndNoTestCounter).strFailAndNoTestFlag = site & " " & strBlockFlagName & " " & strTestFlagName & " NotTest" & " " & strTestFlagCheckInfor
                        lFailAndNoTestCounter = lFailAndNoTestCounter + 1
                        If (strTestFlagCheckInfor = "V" Or (strTestFlagCheckInfor = "S" And (currentJobName Like "ft*"))) Then         ' if mark as "V", we think that it's fail if the flag is clear. but if mark as "C", we think that it's pass if the flag is clear
                           lFailCounter = lFailCounter + 1
                        End If
                    End If  'If FlagStatus = logicTrue Then
               End With      ''  tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(lBlockFlagIdx)
               '======================================================================================================================================================================
               If EnableHVorLV <> False Then              '//EnableHVorLV: true=>HV or LV pass, The test condition is pass
                   If strTestFlagName Like "*HV" Then
                      strTestFlagNameTemp = Replace(strTestFlagName, "HV", "LV")
                   ElseIf strTestFlagName Like "*LV" Then
                      strTestFlagNameTemp = Replace(strTestFlagName, "LV", "HV")
                   End If
                   
                   'search if exsit
                   If strTestFlagNameTemp <> "" Then
                      PreFlagStatus = TheExec.sites.item(site).FlagState(strTestFlagNameTemp)   '//corresponding
                      Dim lFlagTempIdx As Long
                      For lFlagTempIdx = 0 To lFlagIdx - 1
                             If strTestFlagNameTemp = tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(iSiteIdx).astrFlag(lFlagTempIdx).strFlagName Then
                                  strPreTestFlagCheckInfor = tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(iSiteIdx).astrFlag(lFlagTempIdx).strFlagCheckInfor
                                   If (PreFlagStatus = logicFalse) And _
                                          (FlagStatus = logicTrue Or (FlagStatus = logicClear And (strTestFlagCheckInfor = "V" Or (strTestFlagCheckInfor = "S" And (currentJobName Like "ft*"))))) Then     ''' strTestFlagNameTemp is passand strTestFlagName is fail
                                          lFailCounter = lFailCounter - 1
                                          strTestFlagNameTemp = vbNullString
                                   ElseIf (PreFlagStatus = logicClear And strPreTestFlagCheckInfor = "C") And _
                                          (FlagStatus = logicTrue Or (FlagStatus = logicClear And (strTestFlagCheckInfor = "V" Or (strTestFlagCheckInfor = "S" And (currentJobName Like "ft*"))))) Then     ''' strTestFlagNameTemp is pass and strTestFlagName is fail
                                          lFailCounter = lFailCounter - 1
                                          strTestFlagNameTemp = vbNullString
                                   ElseIf (PreFlagStatus = logicTrue) And _
                                          (FlagStatus = logicFalse Or (FlagStatus = logicClear And strTestFlagCheckInfor = "C")) Then      ' '' strTestFlagNameTemp is pass and strTestFlagName is fail
                                          lFailCounter = lFailCounter - 1
                                          strTestFlagNameTemp = vbNullString
                                   ElseIf (PreFlagStatus = logicClear And (strPreTestFlagCheckInfor = "V" Or (strPreTestFlagCheckInfor = "S" And currentJobName Like "ft*"))) And _
                                          (FlagStatus = logicFalse Or (FlagStatus = logicClear And strTestFlagCheckInfor = "C")) Then      ' '' strTestFlagNameTemp is pass and strTestFlagName is fail
                                          lFailCounter = lFailCounter - 1
                                          strTestFlagNameTemp = vbNullString
                                   End If
                                   If strTestFlagCheckInfor <> strPreTestFlagCheckInfor Then
                                           TheExec.Datalog.WriteComment (" The check information of " & strTestFlagNameTemp & " and " & strTestFlagName & " in " & strBlockFlagName & " is different! ")
                                   End If
                                   Exit For
                           End If
                      Next lFlagTempIdx
                   End If  ' If strTestFlagNameTemp<> ""
               End If ' If EnableHVorLV
               '======================================================================================================================================================================
            Next lFlagIdx

            If blIsFlagOrPass <> False And UCase(BlockContext) Like "*FLAGOR-*" Then '' 20160504 webster
                   lFailCounter = 0
                   blIsFlagOrPass = False
            End If

            tyAllBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrSiteVal(site).lFailCounts = lFailCounter  '存 BinOutCalcScanMbistTable 每個block的fail count
            lFailCounter = 0
        Next site
        

        ReDim Preserve tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(lBlockFlagIdx).astrFlag(lFailAndNoTestCounter - 1)
        lFailAndNoTestCounter = 0
  Next lBlockFlagIdx



Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''' Update for TTR  20160504  ---- webster
Public Function BinOutCalc_AssignBinOutFlag(Optional blDebugMode As Boolean = False) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "BinOutCalc_AssignBinOutFlag"

    Dim site As Variant
    Dim iBlockFlagIdx As Integer
    Dim lFlagIdx As Long
    Dim lBinCount As Long
    Dim lBlockCount As Long
    Dim lFailCount As Long
    Dim lBlkFailCount As Long
    Dim lFlagNameCount As Long
    Dim strBlkNam As String
    Dim strBlockNamTemp As String
    Dim astrBlkNam() As String
    Dim strBinFlgNam As String
    
' V: not test logicClear => fail:logicTrue       ; have to be tested       , if flag is clear and setting "V", same with test fail
' C: not test logicClear => pass:logicFalse      ; Do not have to be tested, if flag is clear and setting "C", same with test pass
' "": do not care: not Binning
' S: the same with "V", only use in FT
' tyBinOutInfor.astrBinFlagName(BinFlagNum - 1)       : BinOutTable_CP1/BinOutTable_CP2/BinOutTable_FT1
' tyAllBlockFlagInfor.astrBlockFlag (iBlockFlagIdx)   : BinOutCalcScanMbistTable
' Flag        : And (Flag-CPU_Chain_SA_all)
' Fail        : at least one flag fail
' FlagOr      : Or (FlagOr-CPU_Chain_SA_all)
' Pass        : at least one flag pass
' EnableHVorLV: true=>HV or LV pass, The test condition is pass

    If blDebugMode = True Then   'To print all flag information in datalog
        If SafeArrayGetDim(tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag) > 0 Then
            For iBlockFlagIdx = 0 To UBound(tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag)
                For lFlagIdx = 0 To UBound(tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(iBlockFlagIdx).astrFlag)
                    TheExec.Datalog.WriteComment (tyAllFailAndNoTestBlockFlagInfor.astrBlockFlag(iBlockFlagIdx).astrFlag(lFlagIdx).strFailAndNoTestFlag)
                Next lFlagIdx
            Next iBlockFlagIdx
        Else
            TheExec.Datalog.WriteComment ("===============All tests are pass================")
        End If
    End If

    '' Data missing ,  20160523  ----webster
    If SafeArrayGetDim(tyBinOutInfor.astrBinFlagName) <= 0 Or SafeArrayGetDim(tyAllBlockFlagInfor.astrBlockFlag) <= 0 Then
        TheExec.Datalog.WriteComment ("Read Excel Sheet data again")
        blReadBinOutFlag = False
        Call BinOutCalc_CleanFlagStatus("BinOutCalcScanMbistTable", "BinOutCalcHardIPTable", "BinOutTable", False)    '' read sheet data again, and do not clear test flag status
        Call BinOutCalc_ReadBlockFlagState
        If SafeArrayGetDim(tyBinOutInfor.astrBinFlagName) <= 0 Or SafeArrayGetDim(tyAllBlockFlagInfor.astrBlockFlag) <= 0 Then
            TheExec.Datalog.WriteComment ("<Warning> Data missing, please revalidate to get BinOutCalcTable information in step1 ")
'DataMissing:
            For Each site In TheExec.sites
                TheExec.sites.item(site).FlagState("F_Binning_Error") = logicTrue
                TheExec.sites.item(site).FlagState("F_BIN1") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN2") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN3") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN4") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN5") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN6") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN7") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN8") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN9") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN34") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_GFX_Info_Phantom0") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_GFX_Info_Phantom1") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_GFX_Info_Phantom2") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_Ecore0") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_Pcore0") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_Ecore1") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_Pcore1") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_Ecore2") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_Pcore2") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_WayGroup0") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_WayGroup1") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_WayGroup2") = logicFalse
                TheExec.sites.item(site).FlagState("F_ES_CPU_Info_WayGroup3") = logicFalse
                TheExec.sites.item(site).SortNumber = 999
                TheExec.sites.item(site).BinNumber = 999
                TheExec.sites.item(site).result = tlResultFail
            Next site
            Exit Function       ''end
        End If
    End If
    
    Dim datamiss As Boolean
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////BinOutTable_CP1/BinOutTable_CP2/BinOutTable_FT1 loop
    For Each site In TheExec.sites
        For lBinCount = 0 To UBound(tyBinOutInfor.astrBinFlagName)               '// for BinOutTable_CP1/BinOutTable_CP2/BinOutTable_FT1 loop  ; x loop
            strBinFlgNam = tyBinOutInfor.astrBinFlagName(lBinCount).strBinFlag   '// x index
            
            If strBinFlgNam = "F_BIN1" Then
               Debug.Print "Find F_BIN1"
            End If
            
            If SafeArrayGetDim(tyBinOutInfor.astrBinFlagName(lBinCount).astrBlackName) <= 0 Then
                TheExec.Datalog.WriteComment ("<Warning> Data missing, please revalidate to get BinOutCalcTable information in step2 ")
                'Dim datamiss As Boolean
                datamiss = True
                Exit For
                'GoTo DataMissing
            ElseIf SafeArrayGetDim(tyBinOutInfor.astrBinFlagName(lBinCount).astrBlackName) > 0 Then
                For lBlockCount = 0 To UBound(tyBinOutInfor.astrBinFlagName(lBinCount).astrBlackName)     '//y loop
                    strBlkNam = tyBinOutInfor.astrBinFlagName(lBinCount).astrBlackName(lBlockCount)       '//y value

                    If UCase(strBlkNam) Like "*FLAG*" Then                                                '//do "or" "and" Calculation
                        For iBlockFlagIdx = 0 To UBound(tyAllBlockFlagInfor.astrBlockFlag)
                            strBlockNamTemp = tyAllBlockFlagInfor.astrBlockFlag(iBlockFlagIdx).strBlockFlagName & "-E"   '//add "-E" ; from BinOutCalcScanMbistTable's parameter

                            If InStr(strBlkNam, strBlockNamTemp) Then
                                If tyAllBlockFlagInfor.astrBlockFlag(iBlockFlagIdx).astrSiteVal(site).lFailCounts > 0 Then   'It's block & block ; replace true or false become 1 or o
                                    strBlkNam = Replace(strBlkNam, strBlockNamTemp, 0)
                                Else
                                    strBlkNam = Replace(strBlkNam, strBlockNamTemp, 1)
                                End If
                            End If
                            '' need update
                            If InStr(UCase(strBlkNam), "FLAG") = 0 Then
                                Exit For
                            End If
                        Next iBlockFlagIdx
                                                
                        If UCase(strBlkNam) Like "*FLAG*" Then
                            TheExec.Datalog.WriteComment ("This block name(" & strBlkNam & ") is not found")
                            TheExec.sites.item(site).SortNumber = 999
                            TheExec.sites.item(site).BinNumber = 999
                            TheExec.sites.item(site).result = tlResultFail
                            'Stop
                            'Exit Function
                        End If
                        '--------------------------------------------------------------------------------------------------------------------
                        If IsPassOrFail(strBlkNam) = False Then    '//do "or" "and" Calculation and get pass or fail result
                            lBlkFailCount = lBlkFailCount + 1      'In the bin flag, total fail block flag counts
                            Exit For                               'if lBlkFailCount<>0, it will exit since we think the bin flag is fail
                        End If
                         '--------------------------------------------------------------------------------------------------------------------
                    End If

                Next lBlockCount  'For lBlockCount = 0 To UBound(tyBinOutInfor.astrBinFlagName(lBinCount).astrBlackName)
            End If
            '--------------------------------------------------------------------------------------------------------------------
            If lBlkFailCount = 0 Then                                                    '//this flag is true, mean test pass
                 TheExec.sites.item(site).FlagState(strBinFlgNam) = logicTrue
                 TheExec.Datalog.WriteComment (site & " " & strBinFlgNam & " is Pass")
            Else                                                                         '//this flag is fail, mean test fail
                TheExec.Datalog.WriteComment (site & " " & strBinFlgNam & " is Fail")
                TheExec.sites.item(site).FlagState(strBinFlgNam) = logicFalse
'                If UCase(strBinFlgNam) Like "*F_BIN9*" Then
'                    TheExec.Sites.Item(Site).FlagState(strBinFlgNam) = logicTrue
'                End If
            End If
            '--------------------------------------------------------------------------------------------------------------------
            lBlkFailCount = 0
        Next lBinCount
        If datamiss Then Exit For
    Next site
    
    If datamiss Then
     For Each site In TheExec.sites
                TheExec.sites.item(site).FlagState("F_Binning_Error") = logicTrue
                TheExec.sites.item(site).FlagState("F_BIN1") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN2") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN3") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN4") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN5") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN6") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN7") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN8") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN9") = logicFalse
                TheExec.sites.item(site).FlagState("F_BIN34") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_GFX_Info_Phantom0") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_GFX_Info_Phantom1") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_GFX_Info_Phantom2") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Ecore0") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Pcore0") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Ecore1") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Pcore1") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Ecore2") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Pcore2") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_WayGroup0") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_WayGroup1") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_WayGroup2") = logicFalse
'                TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_WayGroup3") = logicFalse
                TheExec.sites.item(site).SortNumber = 999
                TheExec.sites.item(site).BinNumber = 999
                TheExec.sites.item(site).result = tlResultFail
            Next site
            Exit Function
    End If
    
    

''        For Each Site In TheExec.Sites
''        'Myst For Bin1 ~ Bin7
''        If TheExec.Sites.Item(Site).FlagState("F_BIN1") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin1", 1, False)
''        ElseIf TheExec.Sites.Item(Site).FlagState("F_BIN2") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin2", 1, False)
''        '20160531 fuse Bin1 ~ Bin3 Only
''        'ElseIf TheExec.Sites.Item(Site).FlagState("F_BIN3") = logicTrue And _
''        '        TheExec.Sites.Item(Site).FlagState("F_BIN4") = logicTrue Then
''        '    Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin3", 1, False)
''        '    Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin4", 1, False)
''        ElseIf TheExec.Sites.Item(Site).FlagState("F_BIN3") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin3", 1, False)
''        ElseIf TheExec.Sites.Item(Site).FlagState("F_BIN4") = logicTrue Then
''            GoTo NextSiteJump '20160531 fuse Bin1 ~ Bin3 Only
''            'Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin4", 1, False)
''        ElseIf TheExec.Sites.Item(Site).FlagState("F_BIN5") = logicTrue Then
''            GoTo NextSiteJump '20160531 fuse Bin1 ~ Bin3 Only
''            'Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin5", 1, False)
''        ElseIf TheExec.Sites.Item(Site).FlagState("F_BIN6") = logicTrue Then
''            GoTo NextSiteJump '20160531 fuse Bin1 ~ Bin3 Only
''            'Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin6", 1, False)
''        ElseIf TheExec.Sites.Item(Site).FlagState("F_BIN7") = logicTrue Then
''            GoTo NextSiteJump '20160531 fuse Bin1 ~ Bin3 Only
''            'Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin7", 1, False)
'''        ElseIf TheExec.Sites.Item(Site).FlagState("F_BIN8") = logicTrue Then
'''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_Binning_Bin8", 1, False)
''        End If
        'Myst For each core
        ''Only Keep Ecore & Pcore
'        If TheExec.Sites.Item(Site).FlagState("F_ES_GFX_Info_Phantom0") = logicTrue Then
'            Call auto_eFuse_SetWriteDecimal("UDR", "ES_GFX_Info_Phantom0", 1, False)
'        End If
'        If TheExec.Sites.Item(Site).FlagState("F_ES_GFX_Info_Phantom1") = logicTrue Then
'            Call auto_eFuse_SetWriteDecimal("UDR", "ES_GFX_Info_Phantom1", 1, False)
'        End If
'        If TheExec.Sites.Item(Site).FlagState("F_ES_GFX_Info_Phantom2") = logicTrue Then
'            Call auto_eFuse_SetWriteDecimal("UDR", "ES_GFX_Info_Phantom2", 1, False)
'        End If
''        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Ecore0") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_Ecore0", 1, False)
''        End If
''        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Pcore0") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_Pcore0", 1, False)
''        End If
''        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Ecore1") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_Ecore1", 1, False)
''        End If
''        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Pcore1") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_Pcore1", 1, False)
''        End If
''        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Ecore2") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_Ecore2", 1, False)
''        End If
''        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_Pcore2") = logicTrue Then
''            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_Pcore2", 1, False)
''        End If
'        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_WayGroup0") = logicTrue Then
'            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_WayGroup0", 1, False)
'        End If
'        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_WayGroup1") = logicTrue Then
'            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_WayGroup1", 1, False)
'        End If
'        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_WayGroup2") = logicTrue Then
'            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_WayGroup2", 1, False)
'        End If
'        If TheExec.Sites.Item(Site).FlagState("F_ES_CPU_Info_WayGroup3") = logicTrue Then
'            Call auto_eFuse_SetWriteDecimal("UDR", "ES_CPU_Info_WayGroup3", 1, False)
'        End If
        
''NextSiteJump:
''
''    Next Site

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Private Function IsPassOrFail(ByVal strTemp As String) As Boolean
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "IsPassOrFail"
   ' If InStr(strTemp, "|") Then
        strTemp = Replace(strTemp, "|", "+")
   ' End If
   ' If InStr(strTemp, "&") Then
        strTemp = Replace(strTemp, "&", "*")
  '  End If
  '  If InStr(strTemp, "!") Then
        strTemp = Replace(strTemp, "!1", "0"): strTemp = Replace(strTemp, "!0", "1")
  '  End If
    strTemp = strTemp & "+1"
    If Evaluate(strTemp) - 1 = 0 Then
        IsPassOrFail = False
    Else
        IsPassOrFail = True
    End If
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function




