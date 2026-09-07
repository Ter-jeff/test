Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl
    Set ContextMenu = Application.CommandBars("Cell")
    
    ' Delete the controls first to avoid duplicates.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, before:=1)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "GoBack"
        .FaceId = 59
        .Caption = "GoBack (Ctrl+q)"
        .Tag = "My_Cell_Control_Tag"
    End With
    
    ' Add a separator
    ContextMenu.Controls(2).BeginGroup = True
    Application.MacroOptions Macro:="GoBack", Description:="", ShortcutKey:="q"
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)

    Me.BuiltinDocumentProperties("subject") = Sh.Name

End Sub
