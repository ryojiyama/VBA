Attribute VB_Name = "RightClickAction"
Public Sub CustomizeRightClickMenu()
    Dim menu As CommandBarPopup
    Dim MenuItem As CommandBarButton

    ' Delete if already exists to avoid duplicates
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0

    ' Add custom menu item
    Set menu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
    menu.caption = "Custom Menu"
    
    ' Add a menu item for "UniformizeLineGraphAxes"
    Set MenuItem = menu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .caption = "Uniformize Line Graph Axes"
        .onAction = "UniformizeLineGraphAxes"
        .faceId = 59
    End With
End Sub

Public Sub RemoveRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0
End Sub

