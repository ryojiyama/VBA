Attribute VB_Name = "RightClickAction"
Public Sub CustomizeRightClickMenu()
    Dim Menu As CommandBarPopup
    Dim MenuItem As CommandBarButton

    ' Delete if already exists to avoid duplicates
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0

    ' Add custom menu item
    Set Menu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
    Menu.Caption = "Custom Menu"
    
    ' Add a menu item for "UniformizeLineGraphAxes"
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Uniformize Axes"
        .OnAction = "UniformizeLineGraphAxes"
        .FaceId = 438
    End With

    ' Add a menu item for "InspectionSheet_Make"
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Make InspectionSheets"
        .OnAction = "InspectionSheet_Make"
        .FaceId = 212 ' 358 is an example FaceId, you can change it to any valid FaceId
    End With
    
    ' Add a menu item for "InspectionSheet_Make"
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Delete Copied Sheets"
        .OnAction = "DeleteCopiedSheets"
        .FaceId = 358 ' 358 is an example FaceId, you can change it to any valid FaceId
    End With
End Sub


Public Sub RemoveRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0
End Sub




