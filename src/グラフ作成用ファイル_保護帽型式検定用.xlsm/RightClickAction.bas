Attribute VB_Name = "RightClickAction"

Public Sub CustomizeRightClickMenu()
    Dim Menu As CommandBarPopup
    Dim MenuItem As CommandBarButton

    ' Delete if already exists to avoid duplicates
    On Error Resume Next
    Application.CommandBars("Cell").Controls("チャートの最大値を調整する").Delete
    On Error GoTo 0

    ' Add menu item
    Set Menu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
    Menu.Caption = "チャートの最大値を調整する"
    
    ' Add the macro to the menu item
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "チャートの最大値を調整する"
        .OnAction = "UniformizeLineGraphAxes"
        .FaceId = 59 'Optional: adds a small icon
    End With
End Sub

Public Sub ResetRightClickMenu()
    ' Delete custom menu
    On Error Resume Next
    Application.CommandBars("Cell").Controls("チャートの最大値を調整する").Delete
    On Error GoTo 0
End Sub

