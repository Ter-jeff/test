VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoShmooMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PatternTypes
    initPattern = 1
    plPattern = 2
    None = 0
End Enum
'========== const =========='
Private Const ENABLE_AUTO_SHMOO = "EnableAutoShmoo"
Private Const CURRENT_STATE = "CurrentState"

'========== fields =========='
Private autoShmooItem_ As AutoShmooItem

Private preSelectedSite_ As SiteBoolean
Private curSelectedSite_ As SiteBoolean

Private instanceName_ As String
Private isEnableAutoShmoo_ As Boolean

'========== constructor =========='
Sub Class_Initialize()
    instanceName_ = TheExec.DataManager.instancename ' get ins tance name
    Set autoShmooItem_ = New AutoShmooItem ' for fail item
    
    Set curSelectedSite_ = New SiteBoolean
    isEnableAutoShmoo_ = TheExec.enableWord(ENABLE_AUTO_SHMOO) ' check if enable autoshmoo
End Sub

'========== properties =========='
Public Property Get instancename() As String
    instancename = instanceName_
End Property
Public Property Get IsEnableAutoShmoo() As Boolean
    IsEnableAutoShmoo = isEnableAutoShmoo_
End Property
Public Property Let SetFailingSite(site As Integer)
    curSelectedSite_(site) = True
End Property
Private Property Get IsRunningShmoo() As Boolean
    IsRunningShmoo = TheExec.DevChar.Setups.IsRunning
End Property
Public Property Get ShmooItem() As AutoShmooItem
    Set ShmooItem = autoShmooItem_
End Property

'========== methods =========='
Private Function SaveCurrentSelectedSite()
    ''' save original selected site
    Set preSelectedSite_ = TheExec.sites.Selected
End Function
Private Function RestoreSelectedSite()
    ''' revert back to original selected site
    TheExec.sites.Selected = preSelectedSite_
End Function
Private Function EnableFailingSite()
    ''' enable failing site only
    TheExec.sites.Selected = curSelectedSite_
End Function
Private Function StartShmoo(CharSetupName As String)
    If TheExec.DevChar.Setups.IsRunning = True Then Exit Function
    
    With TheExec.DevChar.Setups(CharSetupName)
        .SaveState CURRENT_STATE
        .Execute False
        .RestoreState CURRENT_STATE
    End With
End Function
Public Function TestInstance()
    If Not isEnableAutoShmoo_ Then Exit Function
    
    'autoShmooItem_.ApplyTimingLevel ' apply timing level
    Dim Pattern As Variant
    
    ''' switch back to vmain setup
    autoShmooUtility.SwitchVmainValt tlDCVSVoltageMain
    ''' set shmoo voltage to valt and setup dssc
    autoShmooItem_.SetShmooCondition autoShmooUtility.GetCurrentCharSetupName
    
    ''' test init patterns
    For Each Pattern In autoShmooItem_.TestinitPatterns
        TheHdw.patterns(Pattern).test pfAlways, 0, tlResultModeDomain
    Next Pattern
    
    ''' switch to valt for mbist block only
    If autoShmooItem_.IsMbistBlock Then autoShmooUtility.SwitchVmainValt tlDCVSVoltageAlt
    ''' test payload patterns
    For Each Pattern In autoShmooItem_.TestPlPatterns
        TheHdw.patterns(Pattern).test pfAlways, 0, tlResultModeDomain
    Next Pattern
    
End Function
Public Function start()
    If Not isEnableAutoShmoo_ Then Exit Function
    
    ''' start to run the shmoo for failing item alone with passing site?? '''
    Dim charSetup As Variant
    
    'SaveCurrentSelectedSite ' save original sites
    'EnableFailingSite ' enable failing sites
    
    If Not autoShmooItem_.IsShmooNeeded Then Exit Function ' no any pattern fail
    
    If Not autoShmooItem_.SetupItem Then Exit Function
    'If Not autoShmooItem_.SetupShmoo Then Exit Function ' setup shmoo and check if valid

    For Each charSetup In autoShmooItem_.CharShmooSetupsOnName.Items
        charSetup.SetShmooTestPattern ' set shmoo test pattern for printing
        If Not IsRunningShmoo Then StartShmoo charSetup.CharSetupName ' start plot shmoo
        charSetup.RestoreShmooTestPattern ' restore shmoo test pattern from global
        autoShmooItem_.CharShmooSetupsOnName.Remove charSetup.CharSetupName ' remove from previous run
    Next charSetup
    
    ''' move test instance inside failing item instead of here
    'TestInstance ' run test instance
    
    'RestoreSelectedSite ' restore original sites
End Function

