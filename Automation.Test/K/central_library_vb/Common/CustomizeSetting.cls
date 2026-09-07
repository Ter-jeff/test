VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomizeSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [20230801][All][James] Add Option Explicit fix Type_Conti
Public Function ReportTname(ProjectType As ProjType, ModuleType As BlockType, Optional PinGroupName As String, Optional ForceIValue As Double) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim TnameSample As String
    Dim instanceName As String
    
    instanceName = UCase(Theexec.DataManager.instanceName)

    If ProjectType = Type_AP Then
        If ModuleType = Type_Conti Then TnameSample = "HAC_<MeasureType>_X_<MeasureCase>_Continuity_<PinName>_<PinGroup>_X_X_<PinSequence>"
    ElseIf ProjectType = Type_LCD Then
        If ModuleType = Type_Conti Then TnameSample = "<OpenShort>_<PinName>_X_X_<ForceI>_X_X_X_X_X_X_X_X_X_<PinSequence>"
    ElseIf ProjectType = Type_RF Then
        If ModuleType = Type_Conti Then TnameSample = "DCTEST_<TestType>_<PinName>_<MeasureCase>_X_X_X_<StartEnd>_NV_VXXX_<PinSequence>"
    End If
    
    If ModuleType = Type_Conti Then ReportTname = CustomizeNameConti(instanceName, TnameSample, ProjectType, PinGroupName, ForceIValue)

                    

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "CustomizeSetting", "ReportTname") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230801][All][James] Add Option Explicit fix not define PinGroupStr
Public Function CustomizeNameConti(InstName As String, InName As String, Proj As ProjType, Optional PinGroup As String, Optional IValue As Double) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim OutName As String
    Dim MeasureType As String
    Dim MeasureCase As String
    Dim ForceI As String
    Dim StartEndCase As String
    Dim TestType As String
    Dim PinGroupStr As String
    
    
    If Proj = Type_AP Then
        MeasureType = "MeasV"
        If InStr(InstName, "POWER_SHORT") <> 0 Then
            MeasureCase = "PS"
            MeasureType = "MeasI"
        ElseIf InStr(InstName, "NEG") <> 0 Then
            MeasureCase = "IONEG"
        ElseIf InStr(InstName, "POS") <> 0 Then
            MeasureCase = "IOPOS"
        Else
        
        End If
        PinGroupStr = Replace(PinGroup, "_", "")
    ElseIf Proj = Type_LCD Then
        TestType = "<OpenShort>"
        If InStr(InstName, "POWER_SHORT") <> 0 Then
            TestType = "PowerShort"
            ForceI = "<ForceI>"
        Else
            ForceI = CStr(IValue * 1000000) & "uA"
        End If
    ElseIf Proj = Type_RF Then
        If InstName Like "*_POS_*" Then
            MeasureCase = "POS"
        ElseIf InstName Like "*_NEG_*" Then
            MeasureCase = "NEG"
        Else
            MeasureCase = "X"
        End If
        TestType = "Conti"
        If InStr(InstName, "POWER_SHORT") <> 0 Then
            TestType = "PowerShort"
        End If
        StartEndCase = "SOF"
        If Name_Flag = 1 Then StartEndCase = "EOF"
    End If
    
    OutName = Replace(InName, "<MeasureType>", MeasureType)
    OutName = Replace(OutName, "<MeasureCase>", MeasureCase)
    OutName = Replace(OutName, "<PinGroup>", PinGroupStr)
    OutName = Replace(OutName, "<ForceI>", ForceI)
    OutName = Replace(OutName, "<StartEnd>", StartEndCase)
    OutName = Replace(OutName, "<OpenShort>", TestType)
    OutName = Replace(OutName, "<TestType>", TestType)

    CustomizeNameConti = OutName
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "CustomizeSetting", "CustomizeNameConti") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Sub SetContiFlagOnelimit(Hlimit As Variant, Llimit As Variant, MeasureInfo As PinData, OpenFlagName As String, ShortFlagName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim site As Variant
    
    For Each site In Theexec.sites
        If MeasureInfo.value(site) > Hlimit Then
            If Hlimit > 0 Then
                Theexec.sites.item(site).FlagState(OpenFlagName) = logicTrue
            Else
                Theexec.sites.item(site).FlagState(ShortFlagName) = logicTrue
            End If
                            
        ElseIf MeasureInfo.value(site) < Llimit Then
            If Llimit > 0 Then
                Theexec.sites.item(site).FlagState(ShortFlagName) = logicTrue
            Else
                Theexec.sites.item(site).FlagState(OpenFlagName) = logicTrue
            End If
                                
        End If
    Next site
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "CustomizeSetting", "SetContiFlagOnelimit") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


' [20230801][All][James] Add Option Explicit fix input currect FlagName
Public Sub SetContiFlagTwoLimit(Hlimit As Variant, Llimit As Variant, MeasureInfo As PinData, FlagName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim site As Variant

    For Each site In Theexec.sites
        If MeasureInfo.value(site) > Hlimit Or MeasureInfo.value(site) < Llimit Then
            Theexec.sites.item(site).FlagState(FlagName) = logicTrue
        End If
    Next site

Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "CustomizeSetting", "SetContiFlagTwoLimit") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub
