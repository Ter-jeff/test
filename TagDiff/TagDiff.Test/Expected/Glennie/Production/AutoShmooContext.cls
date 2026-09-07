VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoShmooContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========== const =========='
'========== fields =========='
Private levelSheet_ As String
Private dcCategory_ As String
Private dcSelector_ As String
Private timeSetSheet_ As String
Private acCategory_ As String
Private acSelector_ As String

'========== constructor =========='
Sub class_initialzie()
    SetCurrentContext TheExec.DataManager.instancename
End Sub
'========== properties =========='
''' get level sheet
Public Property Get levelsheet() As String
    levelsheet = levelSheet_
End Property
''' get dc category
Public Property Get DCCategory() As String
    DCCategory = dcCategory_
End Property
''' get dc selector
Public Property Get DCSelector() As String
    DCSelector = dcSelector_
End Property
''' get time set sheet
Public Property Get TimeSetSheet() As String
    TimeSetSheet = timeSetSheet_
End Property
''' get ac category
Public Property Get ACCategory() As String
    ACCategory = acCategory_
End Property
Public Property Get ACSelector() As String
    ACSelector = acSelector_
End Property
'========== methods =========='
Private Function SetCurrentContext(instance As String)
    ''' dc level
    levelSheet_ = TheExec.TestInstances.item(instance).TimingAndLevels.Sheets.PinLevels
    dcCategory_ = TheExec.Contexts.ActiveSelection.DCCategory
    dcSelector_ = TheExec.Contexts.ActiveSelection.DCSelector
    ''' ac level
    timeSetSheet_ = TheExec.TestInstances.item(instance).TimingAndLevels.Sheets.Timesets
    acCategory_ = TheExec.Contexts.ActiveSelection.ACCategory
    acSelector_ = TheExec.Contexts.ActiveSelection.ACSelector
End Function
Public Function IsSameWithPreviousContext(instance As String) As Boolean
    ''' check if ac/dc level is same as previous item '''
    IsSameWithPreviousContext = _
        levelSheet_ = TheExec.TestInstances.item(instance).TimingAndLevels.Sheets.PinLevels And _
        dcCategory_ = TheExec.Contexts.ActiveSelection.DCCategory And _
        dcSelector_ = TheExec.Contexts.ActiveSelection.DCSelector And _
        timeSetSheet_ = TheExec.TestInstances.item(instance).TimingAndLevels.Sheets.Timesets And _
        acCategory_ = TheExec.Contexts.ActiveSelection.ACCategory And _
        acSelector_ = TheExec.Contexts.ActiveSelection.ACSelector
End Function
Public Function ApplyTimingLevel()
    TheHdw.Digital.ApplyLevelsTiming True, True, True, , , , , _
        levelSheet_, dcCategory_, dcSelector_, _
        timeSetSheet_, acCategory_, acSelector_
End Function
