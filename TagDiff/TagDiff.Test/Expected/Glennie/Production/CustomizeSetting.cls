VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomizeSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function ReportTname(ProjectType As ProjType, moduleType As BlockType, Optional pingroupname As String, Optional ForceIValue As Double) As String

    Dim TnameSample As String
    Dim instancename As String
    
    instancename = UCase(TheExec.DataManager.instancename)

    If ProjectType = Type_AP Then
        If moduleType = Type_Conti Then TnameSample = "HAC_<MeasureType>_X_<MeasureCase>_Continuity_<PinName>_<PinGroup>_X_X_<PinSequence>"
        If moduleType = Type_AutoZ Then TnameSample = "AutoZ_<MeasureType>_X_<MeasureCase>_Continuity_<PinName>_<PinGroup>_X_X_<PinSequence>"
        If moduleType = type_Cre Then TnameSample = "pwr_FVMI_<PinName>"
    ElseIf ProjectType = Type_LCD Then
        If moduleType = TypeConti Then TnameSample = "<OpenShort>_<PinName>_X_X_<ForceI>_X_X_X_X_X_X_X_X_X_<PinSequence>"
    ElseIf ProjectType = Type_RF Then
        If moduleType = TypeConti Then TnameSample = "DCTEST_<TestType>_<PinName>_<MeasureCase>_X_X_X_<StartEnd>_NV_VXXX_<PinSequence>"
    End If
    
    If moduleType = Type_Conti Or Type_AutoZ Then ReportTname = CustomizeNameConti(instancename, TnameSample, ProjectType, pingroupname, ForceIValue)

                    

End Function


Public Function CustomizeNameConti(InstName As String, InName As String, Proj As ProjType, Optional PinGroup As String, Optional IValue As Double) As String

    Dim OutName As String
    Dim MeasureType As String
    Dim measureCase As String
    Dim ForceI As String
    Dim StartEndCase As String
    Dim testType As String
    
    
    If Proj = Type_AP Then
        MeasureType = "MeasV"
        If InStr(InstName, "POWER_SHORT") <> 0 Then
            measureCase = "PS"
            MeasureType = "MeasI"
        ElseIf InStr(InstName, "NEG") <> 0 Then
            measureCase = "IONEG"
        ElseIf InStr(InstName, "POS") <> 0 Then
            measureCase = "IOPOS"
        Else
            measureCase = "<MeasureCase>"
        End If
        PinGroupStr = Replace(PinGroup, "_", "")
    ElseIf Proj = Type_LCD Then
        testType = "<OpenShort>"
        If InStr(InstName, "POWER_SHORT") <> 0 Then
            testType = "PowerShort"
            ForceI = "<ForceI>"
        Else
            ForceI = CStr(IValue * 1000000) & "uA"
        End If
    ElseIf Proj = Type_RF Then
        If InstName Like "*_POS_*" Then
            measureCase = "POS"
        ElseIf InstName Like "*_NEG_*" Then
            measureCase = "NEG"
        Else
            measureCase = "X"
        End If
        testType = "Conti"
        If InStr(InstName, "POWER_SHORT") <> 0 Then
            testType = "PowerShort"
        End If
        StartEndCase = "SOF"
        If Name_Flag = 1 Then StartEndCase = "EOF"
    End If
    
    OutName = Replace(InName, "<MeasureType>", MeasureType)
    OutName = Replace(OutName, "<MeasureCase>", measureCase)
    OutName = Replace(OutName, "<PinGroup>", PinGroupStr)
    OutName = Replace(OutName, "<ForceI>", ForceI)
    OutName = Replace(OutName, "<StartEnd>", StartEndCase)
    OutName = Replace(OutName, "<OpenShort>", testType)
    OutName = Replace(OutName, "<TestType>", testType)

    CustomizeNameConti = OutName
End Function


Public Sub SetContiFlagOnelimit(Hlimit As Variant, Llimit As Variant, MeasureInfo As PinData, OpenFlagName As String, ShortFlagName As String)

    Dim site As Variant
    
    For Each site In TheExec.sites
        If MeasureInfo.value(site) > Hlimit Then
            If Hlimit > 0 Then
                TheExec.sites.item(site).FlagState(OpenFlagName) = logicTrue
            Else
                TheExec.sites.item(site).FlagState(ShortFlagName) = logicTrue
            End If
                            
        ElseIf MeasureInfo.value(site) < Llimit Then
            If Llimit > 0 Then
                TheExec.sites.item(site).FlagState(ShortFlagName) = logicTrue
            Else
                TheExec.sites.item(site).FlagState(OpenFlagName) = logicTrue
            End If
                                
        End If
    Next site
End Sub


Public Sub SetContiFlagTwoLimit(Hlimit As Variant, Llimit As Variant, MeasureInfo As PinData, flagName As String)

    Dim site As Variant

    For Each site In TheExec.sites
        If MeasureInfo.value(site) > hiLimit Or MeasureInfo.value(site) < LowLimit Then
            TheExec.sites.item(site).FlagState("FlagName") = logicTrue
        End If
    Next site

End Sub

