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
        .Caption = "Uniformize Line Graph Axes"
        .OnAction = "UniformizeLineGraphAxes"
        .FaceId = 59
    End With
End Sub

Public Sub RemoveRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0
End Sub

