Attribute VB_Name = "RightClickAction"

'Public Sub CustomizeRightClickMenu()
'    Dim Menu As CommandBarPopup
'    Dim MenuItem As CommandBarButton
'
'    ' Delete if already exists to avoid duplicates
'    On Error Resume Next
'    Application.CommandBars("Cell").Controls("Custom Menu").Delete
'    On Error GoTo 0
'
'    ' Add custom menu item
'    Set Menu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'    Menu.Caption = "Custom Menu"
'
'    ' Add a menu item for "Create ID"
'    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
'    With MenuItem
'        .Caption = "Create ID"
'        .OnAction = "ShowFormSheetName"
'        .FaceId = 438
'    End With
'
'    ' Add a menu item for "Sync Spec Sheet"
'    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
'    With MenuItem
'        .Caption = "Sync Spec Sheet"
'        .OnAction = "SyncSpecSheetToLogHel"
'        .FaceId = 212
'    End With
'End Sub
'
'
'
'Public Sub RemoveRightClickMenu()
'    On Error Resume Next
'    Application.CommandBars("Cell").Controls("Custom Menu").Delete
'    On Error GoTo 0
'End Sub


'Sub ShowFormSheetName()
'    Form_SheetName.Show
'End Sub
