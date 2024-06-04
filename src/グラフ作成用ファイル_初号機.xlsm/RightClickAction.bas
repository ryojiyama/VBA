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
    
    ' Add a menu item for "DeleteAllChartsInActiveSheet"
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Delete All Charts in Active Sheet"
        .OnAction = "DeleteAllChartsInActiveSheet"
        .FaceId = 60 ' Change the FaceId as required
    End With
    
    ' Add a menu item to show "UserForm1"
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Show User Form 1"
        .OnAction = "ShowUserForm1"
        .FaceId = 61 ' Change the FaceId as required
    End With
End Sub

' Additional Sub to show UserForm1
Sub ShowUserForm1()
    UserForm1.Show
End Sub


Public Sub ResetRightClickMenu()
    ' Delete custom menu
    On Error Resume Next
    Application.CommandBars("Cell").Controls("UniformizeLineGraphAxes").Delete
    On Error GoTo 0
End Sub

